use calamine::Reader;
use calamine::Xlsx;
use calamine::open_workbook;
use dotenv::dotenv;
use eframe::egui;
use std::env;
use std::error::Error;
use std::path::Path;
use std::thread;
use std::time::Duration;
use thirtyfour::prelude::*;
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

    let driver = WebDriver::new("http://localhost:9515", caps).await?;
    let domain = "http://www.dongkun.com";
    let base_url = format!("{}/ko/sub/product", domain);
    
    // 먼저 카테고리 목록 페이지로 이동
    let list_url = format!("{}/list.asp?s_cate=10", base_url);
    driver.goto(&list_url).await?;
    tokio::time::sleep(Duration::from_secs(2)).await;

    // 모든 카테고리 링크 수집
    let categories = driver
        .find_elements(By::Css(".category_li > ul > li > a"))
        .await?;

    let mut category_info = Vec::new();
    for category in &categories {
        if let (Ok(Some(href)), Ok(name)) = (category.get_attribute("href").await, category.text().await) {
            // href가 상대 경로로 시작하면 도메인을 앞에 추가
            let full_href = if href.starts_with("/") {
                format!("{}{}", domain, href)
            } else {
                href.to_string()
            };
            category_info.push((full_href, name.to_string()));
        }
    }


    // 각 카테고리 순회
    for (index, (category_href, category_name)) in category_info.iter().enumerate() {
        println!("\n=== 카테고리 {}/{}: {} ===", index + 1, category_info.len(), category_name);
        
        // 카테고리 페이지로 이동
        driver.goto(category_href).await?;
        tokio::time::sleep(Duration::from_secs(1)).await;

        // 현재 카테고리의 제품 링크 수집
        let product_links = driver
            .find_elements(By::Css(
                "ul.clearfix > li > a[href*='idx=']:not([href$='view.asp'])",
            ))
            .await?;

        println!("카테고리 내 제품 수: {}", product_links.len());

        // 각 제품의 href 속성에서 파라미터 추출
        let mut product_params = Vec::new();
        for product in &product_links {
            if let Ok(Some(href)) = product.get_attribute("href").await {
                let href = href.to_string();
                if let Some(query) = href.split("?").nth(1) {
                    // 제품명도 함께 저장
                    if let Ok(product_elem) = product.find_element(By::Css("div.txt > p")).await {
                        if let Ok(product_name) = product_elem.text().await {
                            product_params.push((query.to_string(), product_name.to_string()));
                        }
                    }
                }
            }
        }

        // 각 제품 페이지 방문
        for (prod_index, (params, product_name)) in product_params.iter().enumerate() {
            println!("\n제품 {}/{}: {}", prod_index + 1, product_params.len(), product_name);
            
            // 전체 URL 구성
            let full_url = format!("{}/view.asp?{}", base_url, params);
            println!("방문 URL: {}", full_url);
            
            // 제품 상세 페이지로 이동
            driver.goto(&full_url).await?;
            tokio::time::sleep(Duration::from_secs(1)).await;

            // 카테고리 리스트로 돌아가기
            driver.goto(category_href).await?;
            tokio::time::sleep(Duration::from_secs(1)).await;
        }
    }

    driver.quit().await?;
    Ok(())
}