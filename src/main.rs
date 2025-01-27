use calamine::Reader;
use calamine::Xlsx;
use calamine::open_workbook;
use dotenv::dotenv;
use eframe::egui;
use std::env;
use std::error::Error;
use std::fs;
use std::path::Path;
use std::thread;
use std::time::Duration;
use thirtyfour::prelude::*;
use xlsxwriter::*;
use tokio::process::Command;
use tokio::time::sleep; // Path를 사용하여 디렉토리 경로 조작
const WINDOW_DRIVER: &str = "./driver/chromedriver.exe";
const MAC_DRIVER: &str = "./driver/chromedriver";

const TYPE: &str = "MAC";
fn main() -> eframe::Result<()> {
    // if let Ok(exe_path) = std::env::current_exe() {
    //     if let Some(exe_dir) = exe_path.parent() {
    //         std::env::set_current_dir(exe_dir).ok();
    //     }
    // }
    dotenv().ok(); // .env 파일 로드

    let app = MyApp::default();
    let options = eframe::NativeOptions::default();
    eframe::run_native("jang sung jin", options, Box::new(|_cc| Ok(Box::new(app))))
}

struct MyApp {
    shown1: bool,
    shown2: bool,
    web_status: String,
}
async fn start_chromedriver() -> Result<(), Box<dyn Error + Send + Sync>> {
    let chromedriver_path = if TYPE == "WINDOW" {
        WINDOW_DRIVER
    } else {
        MAC_DRIVER
    };
    // 기존 크롬드라이버 프로세스 종료 추가
    Command::new("killall").arg("chromedriver").spawn().ok();
    sleep(Duration::from_secs(1)).await;

    Command::new(chromedriver_path).arg("--port=9515").spawn()?;
    sleep(Duration::from_secs(5)).await;
    Ok(())
}

pub async fn example() -> Result<(), Box<dyn Error + Send + Sync>> {
    start_chromedriver().await?;

    let mut caps = DesiredCapabilities::chrome();
    caps.set_no_sandbox()?;

    let driver = WebDriver::new("http://localhost:9515", caps).await?;
    driver
        .goto("https://ikor250113s.mycafe24.com/bbs/board.php?bo_table=product")
        .await?;

    match driver.find_element(By::ClassName("adminLi")).await {
        Ok(_) => {
            println!("이미 관리자 페이지입니다");
        }
        Err(_) => {
            driver.goto("https://ikor250113s.mycafe24.com/adm").await?;
            let admin_id = env::var("ADMIN_ID").expect("ADMIN_ID must be set");
            let admin_pw = env::var("ADMIN_PW").expect("ADMIN_PW must be set");

            if let Ok(alert) = driver.get_alert_text().await {
                driver.accept_alert().await?; // 알림창 확인 클릭
            }

            driver
                .find_element(By::Name("mb_id"))
                .await?
                .send_keys(admin_id)
                .await?;

            driver
                .find_element(By::Name("mb_password"))
                .await?
                .send_keys(admin_pw)
                .await?;
            // println!("{}", admin_id);
            driver
                .find_element(By::ClassName("btn_submit"))
                .await?
                .click()
                .await?;

            sleep(Duration::from_secs(1)).await; // 로그인 처리 대기
            driver
                .goto("https://ikor250113s.mycafe24.com/bbs/board.php?bo_table=product")
                .await?;
            sleep(Duration::from_secs(1)).await; // 로그인 처리 대기

            driver
                .find_element(By::Css("a.btn_b01.btn"))
                .await?
                .click()
                .await?;

            // Excel 파일 읽기
            if let Ok(current_dir) = env::current_dir() {
                println!("Current directory: {}", current_dir.display());
            }
            let path = "./test.xlsx"; // 현재 디렉토리의 파일 지정

            let mut workbook: Xlsx<_> = open_workbook(path)?;
            if let Some(Ok(range)) = workbook.worksheet_range_at(0) {
                println!("총 제품 수: {}", range.height() - 1);
                for row_idx in 1..range.height() {
                    println!(
                        "제품 {}: {:?}",
                        row_idx,
                        range.get_value((row_idx as u32, 0))
                    );
                }
            }
        }
    }

    sleep(Duration::from_secs(3)).await;

    driver.quit().await?;

    Ok(())
}

impl Default for MyApp {
    fn default() -> Self {
        Self {
            shown1: false,
            shown2: false,
            web_status: "Ready".to_string(),
        }
    }
}

impl eframe::App for MyApp {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        egui::CentralPanel::default().show(ctx, |ui| {
            if ui.button("Start Web Automation").clicked() {
                self.shown1 = !self.shown1;
                if self.shown1 {
                    self.web_status = "Starting...".to_string();
                    // 새로운 스레드에서 웹 자동화 실행
                    thread::spawn(|| {
                        let rt = tokio::runtime::Runtime::new().unwrap();
                        rt.block_on(async {
                            if let Err(e) = example().await {
                                println!("Error: {:?}", e);
                            }
                        });
                    });
                }
            }
            ui.label(&self.web_status);

            ui.add_space(20.0);

            if ui.button("Go to Dongkun").clicked() {
                self.shown2 = !self.shown2;
                if self.shown2 {
                    thread::spawn(|| {
                        let rt = tokio::runtime::Runtime::new().unwrap();
                        rt.block_on(async {
                            if let Err(e) = dongkun_example().await {
                                println!("Error: {:?}", e);
                            }
                        });
                    });
                }
            }
        });
    }
}

pub async fn dongkun_example() -> Result<(), Box<dyn Error + Send + Sync>> {
    start_chromedriver().await?;
    let mut caps = DesiredCapabilities::chrome();
    caps.set_no_sandbox()?;

    // Excel 워크북 생성
    let workbook = Workbook::new("products.xlsx")?;
    let mut sheet = workbook.add_worksheet(None)?;

    // 헤더 작성
    sheet.write_string(0, 0, "분류", None)?;
    sheet.write_string(0, 1, "제목", None)?;
    sheet.write_string(0, 2, "제품특징", None)?;
    sheet.write_string(0, 3, "사용장소", None)?;
    sheet.write_string(0, 4, "사진1", None)?;
    sheet.write_string(0, 5, "사진2", None)?;

    let mut row = 1; // 데이터는 1행부터 시작

    let driver = WebDriver::new("http://localhost:9515", caps).await?;
    let domain = "http://www.dongkun.com";
    let base_url = format!("{}/ko/sub/product", domain);

    let main_url = format!("{}/list.asp", base_url);
    driver.goto(&main_url).await?;
    tokio::time::sleep(Duration::from_secs(1)).await;

    let main_categories = driver
        .find_elements(By::Css(".depth2.menu2 > li > a"))
        .await?;

    let mut main_category_info = Vec::new();
    for category in &main_categories {
        if let (Ok(Some(href)), Ok(name)) = (category.get_attribute("href").await, category.text().await) {
            let full_href = if href.starts_with("/") {
                format!("{}{}", domain, href)
            } else {
                href.to_string()
            };
            main_category_info.push((full_href, name.to_string()));
        }
    }

    for (main_index, (main_href, main_name)) in main_category_info.iter().enumerate() {
        println!("\n=== 메인 카테고리 {}/{}: {} ===", main_index + 1, main_category_info.len(), main_name);
        
        driver.goto(main_href).await?;
        tokio::time::sleep(Duration::from_secs(1)).await;

        let sub_categories = driver
            .find_elements(By::Css(".category_li > ul > li > a"))
            .await?;

        let mut sub_category_info = Vec::new();
        for category in &sub_categories {
            if let (Ok(Some(href)), Ok(name)) = (category.get_attribute("href").await, category.text().await) {
                let full_href = if href.starts_with("/") {
                    format!("{}{}", domain, href)
                } else {
                    href.to_string()
                };
                sub_category_info.push((full_href, name.to_string()));
            }
        }

        for (sub_index, (sub_href, sub_name)) in sub_category_info.iter().enumerate() {
            println!("\n-- 서브 카테고리 {}/{}: {} --", sub_index + 1, sub_category_info.len(), sub_name);
            
            driver.goto(sub_href).await?;
            tokio::time::sleep(Duration::from_secs(1)).await;

            let products = driver
                .find_elements(By::Css("ul.clearfix > li > a"))
                .await?;

            let mut product_info = Vec::new();
            for product in &products {
                if let Ok(Some(href)) = product.get_attribute("href").await {
                    if let Some(query) = href.split("?").nth(1) {
                        if let Ok(product_elem) = product.find_element(By::Css("div.txt > p")).await {
                            if let Ok(product_name) = product_elem.text().await {
                                let full_href = format!("{}/view.asp?{}", base_url, query);
                                product_info.push((full_href, product_name.to_string()));
                            }
                        }
                    }
                }
            }

            println!("제품 수: {}", product_info.len());

            for (prod_index, (prod_href, prod_name)) in product_info.iter().enumerate() {
                println!("제품 {}/{}: {}", prod_index + 1, product_info.len(), prod_name);
                println!("URL: {}", prod_href);
                
                driver.goto(prod_href).await?;
                tokio::time::sleep(Duration::from_secs(1)).await;

                // 상세 페이지에서 정보 수집
                let mut category = String::new();
                let mut name = String::new();
                if let Ok(title_div) = driver.find_element(By::Css(".title")).await {
                    if let Ok(span) = title_div.find_element(By::Tag("span")).await {
                        if let Ok(text) = span.text().await {
                            category = text;
                        }
                    }
                    if let Ok(strong) = title_div.find_element(By::Tag("strong")).await {
                        if let Ok(text) = strong.text().await {
                            name = text;
                        }
                    }
                }

                let mut features = String::new();
                let mut usage = String::new();
                if let Ok(info_div) = driver.find_element(By::Css(".info")).await {
                    let dls = info_div.find_elements(By::Tag("dl")).await?;
                    for dl in dls {
                        if let Ok(dt) = dl.find_element(By::Tag("dt")).await {
                            if let Ok(text) = dt.text().await {
                                if text.contains("제품특징") {
                                    if let Ok(dd) = dl.find_element(By::Tag("dd")).await {
                                        if let Ok(feature_text) = dd.text().await {
                                            features = feature_text;
                                        }
                                    }
                                } else if text.contains("제품 사용장소") {
                                    if let Ok(dd) = dl.find_element(By::Tag("dd")).await {
                                        if let Ok(usage_text) = dd.text().await {
                                            usage = usage_text;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                // 이미지 다운로드
                let img_folder = format!("downloads/{}/{}", category, name);
                fs::create_dir_all(&img_folder)?;

                let mut image1_path = String::new();
                let mut image2_path = String::new();

                // 첫 번째 이미지
                if let Ok(main_img) = driver.find_element(By::Css(".img_box .img img")).await {
                    if let Ok(Some(src)) = main_img.get_attribute("src").await {
                        let full_url = format!("{}{}", domain, src);
                        let file_name = src.split('/').last().unwrap_or("image1.jpg");
                        let save_path = format!("{}/1_{}", img_folder, file_name);
                        
                        download_image(&full_url, &save_path).await?;
                        image1_path = save_path;
                    }
                }

                // 두 번째 이미지
                if let Ok(detail_img) = driver.find_element(By::Css(".detail .txt_area img")).await {
                    if let Ok(Some(src)) = detail_img.get_attribute("src").await {
                        let full_url = format!("{}{}", domain, src);
                        let file_name = src.split('/').last().unwrap_or("image2.jpg");
                        let save_path = format!("{}/2_{}", img_folder, file_name);
                        
                        download_image(&full_url, &save_path).await?;
                        image2_path = save_path;
                    }
                }

                // 엑셀에 데이터 저장
                sheet.write_string(row, 0, &category, None)?;
                sheet.write_string(row, 1, &name, None)?;
                sheet.write_string(row, 2, &features, None)?;
                sheet.write_string(row, 3, &usage, None)?;
                sheet.write_string(row, 4, &image1_path, None)?;
                sheet.write_string(row, 5, &image2_path, None)?;

                row += 1;

                driver.goto(sub_href).await?;
                tokio::time::sleep(Duration::from_secs(1)).await;
            }
        }
    }
    workbook.close()?;

    driver.quit().await?;
    Ok(())
}

// fn save_to_excel(product: &ProductInfo, file_path: &str) -> Result<(), Box<dyn Error>> {
//     let mut workbook = if Path::new(file_path).exists() {
//         Workbook::open(file_path)?
//     } else {
//         Workbook::create(file_path)?
//     };

//     let sheet = workbook.worksheet_mut("Products")?;

//     // 헤더 추가 (첫 번째 행이 비어있는 경우)
//     if sheet.is_empty() {
//         sheet.write_row(&["분류", "제목", "제품특징", "사용장소", "사진1", "사진2"])?;
//     }

//     // 데이터 추가
//     let row = sheet.rows() + 1;
//     sheet.write_row_at(row, &[
//         &product.category,
//         &product.name,
//         &product.features,
//         &product.usage,
//         &product.image1,
//         &product.image2,
//     ])?;

//     workbook.save(file_path)?;
//     Ok(())
// }
async fn download_image(url: &str, path: &str) -> Result<(), Box<dyn Error + Send + Sync>> {
    let response = reqwest::get(url).await?;
    let bytes = response.bytes().await?;
    fs::write(path, bytes)?;
    Ok(())
}
