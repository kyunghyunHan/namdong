use calamine::Reader;
use calamine::Xlsx;
use calamine::open_workbook;
use dotenv::dotenv;
use eframe::egui;
use std::env;
use std::error::Error;
use std::fs;
use std::thread;

use std::time::Duration;
use thirtyfour::prelude::*;
use tokio::process::Command;
use tokio::time::sleep;
use xlsxwriter::*; // Path를 사용하여 디렉토리 경로 조작
const WINDOW_DRIVER: &str = "./driver/chromedriver.exe";
const MAC_DRIVER: &str = "./driver/chromedriver";

const TYPE: &str = "MAC";

const SITE_ADRESS: &str = "http://namdongfan.com";
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
        .goto(format!("{SITE_ADRESS}/bbs/board.php?bo_table=product"))
        .await?;

    match driver.find_element(By::ClassName("adminLi")).await {
        Ok(_) => {
            println!("이미 관리자 페이지입니다");
        }
        Err(_) => {
            driver.goto(format!("{SITE_ADRESS}/adm")).await?;
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

            tokio::time::sleep(Duration::from_millis(400)).await;
            driver
                .goto(format!("{SITE_ADRESS}/bbs/board.php?bo_table=product"))
                .await?;
            sleep(Duration::from_secs(1)).await; // 로그인 처리 대기

            driver
                .find_element(By::Css("a.btn_b01.btn"))
                .await?
                .click()
                .await?;

            if let Ok(current_dir) = env::current_dir() {
                println!("Current directory: {}", current_dir.display());
            }

            let path = "./data/products.xlsx"; // 현재 디렉토리의 파일 지정

            let mut workbook: Xlsx<_> = open_workbook(path)?;

            if let Some(Ok(range)) = workbook.worksheet_range_at(0) {
                // 각 제품에 대해 처리

                for row_idx in 1..range.height() {
                    // 카테고리 선택
                    let category = range
                        .get_value((row_idx as u32, 0)) // 분류 컬럼
                        .map(|v| v.to_string())
                        .unwrap_or_default();

                    let sub_category = range
                        .get_value((row_idx as u32, 1)) // 분류 컬럼
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    // 카테고리 선택 - select 요소의 options를 찾아서 매칭되는 텍스트의 option을 선택
                    let select = driver.find_element(By::Id("ca_name")).await?;
                    let options = select.find_elements(By::Tag("option")).await?;

                    for option in options {
                        let option_text = option.text().await?;
                        println!("{}", option_text);

                        if option_text == category {
                            option.click().await?;
                            break;
                        }
                    }
                    tokio::time::sleep(Duration::from_millis(400)).await;
                    let select = driver.find_element(By::Id("wr_1")).await?;
                    let options = select.find_elements(By::Tag("option")).await?;
                    for option in options {
                        let option_text = option.text().await?;
                        println!("{}", option_text);

                        if option_text == sub_category {
                            option.click().await?;
                            break;
                        }
                    }

                    // 제목 입력
                    let title = range
                        .get_value((row_idx as u32, 2)) // 제목 컬럼
                        .map(|v| v.to_string())
                        .unwrap_or_default();

                    driver
                        .find_element(By::Id("wr_subject"))
                        .await?
                        .send_keys(&title)
                        .await?;

                    // 제품 특징 입력
                    let features = range
                        .get_value((row_idx as u32, 3)) // 특징 컬럼
                        .map(|v| v.to_string())
                        .unwrap_or_default();

                    driver
                        .find_element(By::Id("wr_8"))
                        .await?
                        .send_keys(&features)
                        .await?;

                    // 사용장소 입력
                    let usage = range
                        .get_value((row_idx as u32, 4)) // 사용장소 컬럼
                        .map(|v| v.to_string())
                        .unwrap_or_default();

                    driver
                        .find_element(By::Id("wr_9"))
                        .await?
                        .send_keys(&usage)
                        .await?;
                    let thumbnail_path = range
                        .get_value((row_idx as u32, 5)) // 썸네일 경로 컬럼
                        .map(|v| v.to_string())
                        .unwrap_or_default();

                    if !thumbnail_path.is_empty() {
                        // 상대 경로를 절대 경로로 변환
                        let absolute_path = std::env::current_dir()?.join(&thumbnail_path);
                        println!("업로드할 썸네일 경로: {}", absolute_path.display());

                        // 파일 input 요소를 찾고 파일 경로 전송
                        let file_input = driver.find_element(By::Id("bf_file_1")).await?;

                        // 파일 경로를 문자열로 변환하고 send_keys 실행
                        if let Some(path_str) = absolute_path.to_str() {
                            file_input.send_keys(path_str).await?;
                            // 업로드 후 잠시 대기
                            tokio::time::sleep(Duration::from_secs(1)).await;
                        } else {
                            println!("파일 경로를 문자열로 변환할 수 없습니다");
                        }
                    }

                    let iframe = driver
.find_element(By::Css("iframe[src='http://namdongfan.com/plugin/editor/smarteditor2/SmartEditor2Skin.html']"))
.await?;

                    println!("Found iframe");

                    // iframe으로 전환
                    driver.switch_to().frame_element(&iframe).await?;

                    // iframe 로드를 위해 대기
                    tokio::time::sleep(Duration::from_secs(1)).await;

                    println!("Switched to iframe");

                    // 사진 버튼 찾기
                    let photo_btn = match driver.find_element(By::ClassName("se2_photo")).await {
                        Ok(btn) => btn,
                        Err(_) => {
                            println!("Trying alternative selector...");
                            driver.find_element(By::Css("button.se2_photo")).await?
                        }
                    };

                    println!("Found photo button, clicking...");
                    photo_btn.click().await?;

                    // 사진 버튼 찾기

                    photo_btn.click().await?;
                    let image_path = range
                        .get_value((row_idx as u32, 6)) // 이미지 경로 컬럼
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                    let handles = driver.window_handles().await?;
                    let main_handle = handles.first().unwrap().clone();
                    if !image_path.is_empty() {
                        // 팝업창이 뜰 때까지 대기
                        tokio::time::sleep(Duration::from_secs(1)).await;

                        // 팝업 창으로 전환
                        let handles = driver.window_handles().await?;
                        let popup_handle = handles.last().unwrap();
                        driver.switch_to().window(popup_handle.clone()).await?;

                        // 파일 선택 버튼 찾기
                        let file_select_button = driver
                            .find_element(By::Css("span.fileinput-button"))
                            .await?;

                        // 파일 선택 버튼 클릭
                        file_select_button.click().await?;

                        // 파일 선택 대화상자가 열릴 때까지 대기
                        tokio::time::sleep(Duration::from_millis(500)).await;

                        // 파일 경로 전송
                        let absolute_path = std::env::current_dir()?.join(&image_path);
                        println!("Uploading image from path: {}", absolute_path.display());

                        if let Some(path_str) = absolute_path.to_str() {
                            // 파일 업로드 input 요소 찾기
                            let file_input = driver.find_element(By::Id("fileupload")).await?;
                            println!("{}", path_str);
                            // 파일 경로 입력
                            file_input.send_keys(path_str).await?;

                            // 파일 선택 후 잠시 대기
                            tokio::time::sleep(Duration::from_secs(1)).await;

                            // 등록 버튼 클릭
                            let upload_btn =
                                driver.find_element(By::Id("img_upload_submit")).await?;
                            upload_btn.click().await?;
                        }
                        tokio::time::sleep(Duration::from_millis(1000)).await;
                    } else {
                        // 이미지가 없는 경우 6개의 공백 문자를 가진 p 태그 삽입
                        let script = r#"
    var p = document.createElement('p');
    p.innerHTML = '&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;';
    document.body.appendChild(p);
"#;
                        driver.execute(script, vec![]).await?; // execute_script 대신 execute 사용, 빈 벡터 전달
                    }
                    // 저장해둔 핸들을 사용하여 기존 창으로 전환
                    driver.switch_to().window(main_handle).await?;

                    // 제품 등록 버튼 클릭
                    driver
                        .find_element(By::Id("btn_submit"))
                        .await?
                        .click()
                        .await?;

                    // 다음 제품 입력을 위해 대기
                    tokio::time::sleep(Duration::from_millis(400)).await;
                    driver
                        .goto(format!("{SITE_ADRESS}/bbs/write.php?bo_table=product"))
                        .await?;
                    tokio::time::sleep(Duration::from_millis(400)).await;
                }
            }
        }
    }

    sleep(Duration::from_secs(1)).await;

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
            if ui.button("Upload Data").clicked() {
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

            if ui.button("Saving Data").clicked() {
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
    println!("크롤링 시작...");
    start_chromedriver().await?;
    let mut caps = DesiredCapabilities::chrome();
    caps.set_no_sandbox()?;

    fs::create_dir_all("./data")?;
    let workbook = Workbook::new("./data/products.xlsx")?;
    let mut sheet = workbook.add_worksheet(None)?;

    // 헤더 작성
    sheet.write_string(0, 0, "대분류", None)?;
    sheet.write_string(0, 1, "세부분류", None)?;
    sheet.write_string(0, 2, "제목", None)?;
    sheet.write_string(0, 3, "제품특징", None)?;
    sheet.write_string(0, 4, "사용장소", None)?;
    sheet.write_string(0, 5, "사진1", None)?;
    sheet.write_string(0, 6, "사진2", None)?;

    let mut row = 1;

    let driver = WebDriver::new("http://localhost:9515", caps).await?;
    let domain = "http://www.dongkun.com";
    let base_url = format!("{}/ko/sub/product", domain);
    let main_url = format!("{}/list.asp", base_url);

    // 메인 페이지 로드
    driver.goto(&main_url).await?;
    tokio::time::sleep(Duration::from_millis(200)).await;

    let main_categories = driver
        .find_elements(By::Css(".depth2.menu2 > li > a"))
        .await?;

    let mut main_category_info = Vec::new();
    for category in &main_categories {
        match (category.get_attribute("href").await, category.text().await) {
            (Ok(Some(href)), Ok(name)) => {
                let full_href = if href.starts_with("/") {
                    format!("{}{}", domain, href)
                } else {
                    href.to_string()
                };
                println!("메인 카테고리 발견: {}", name);
                main_category_info.push((full_href, name.to_string()));
            }
            _ => {
                println!("카테고리 정보 추출 실패");
                continue;
            }
        }
    }

    for (main_index, (main_href, main_name)) in main_category_info.iter().enumerate() {
        println!(
            "\n=== 메인 카테고리 {}/{}: {} ===",
            main_index + 1,
            main_category_info.len(),
            main_name
        );

        driver.goto(main_href).await?;
        tokio::time::sleep(Duration::from_millis(200)).await;

        let names = match driver.find_element(By::Css(".pageTit h4")).await {
            Ok(element) => match element.text().await {
                Ok(text) => text,
                Err(_) => {
                    println!("타이틀 텍스트 추출 실패");
                    continue;
                }
            },
            Err(_) => {
                println!("타이틀 요소를 찾을 수 없음");
                continue;
            }
        };

        // 서브 카테고리 확인
        let has_sub_categories = driver.find_element(By::Css(".category_li")).await.is_ok();

        let mut sub_category_info = Vec::new();

        if has_sub_categories {
            // 서브 카테고리가 있는 경우
            if let Ok(sub_categories) = driver
                .find_elements(By::Css(".category_li > ul > li > a"))
                .await
            {
                for category in &sub_categories {
                    if let (Ok(Some(href)), Ok(name)) =
                        (category.get_attribute("href").await, category.text().await)
                    {
                        let full_href = if href.starts_with("/") {
                            format!("{}{}", domain, href)
                        } else {
                            href.to_string()
                        };
                        sub_category_info.push((full_href, name.to_string()));
                    }
                }
            }
        }

        // 서브 카테고리가 없거나 찾지 못한 경우, 대분류 URL을 그대로 사용
        if sub_category_info.is_empty() {
            println!("서브 카테고리 없음, 대분류로 처리: {}", names);
            sub_category_info.push((main_href.clone(), names.clone()));
        }

        for (sub_index, (sub_href, sub_name)) in sub_category_info.iter().enumerate() {
            println!(
                "\n-- 서브 카테고리 {}/{}: {} --",
                sub_index + 1,
                sub_category_info.len(),
                sub_name
            );

            driver.goto(sub_href).await?;
            tokio::time::sleep(Duration::from_millis(200)).await;

            let products = match driver.find_elements(By::Css("ul.clearfix > li > a")).await {
                Ok(elements) => elements,
                Err(e) => {
                    println!("제품 목록을 찾을 수 없음: {:?}", e);
                    continue;
                }
            };

            let mut product_info = Vec::new();
            for product in &products {
                if let Ok(Some(href)) = product.get_attribute("href").await {
                    if let Some(query) = href.split("?").nth(1) {
                        if let Ok(product_elem) = product.find_element(By::Css("div.txt > p")).await
                        {
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
                println!(
                    "제품 {}/{}: {}",
                    prod_index + 1,
                    product_info.len(),
                    prod_name
                );
                println!("URL: {}", prod_href);

                driver.goto(prod_href).await?;
                tokio::time::sleep(Duration::from_millis(200)).await;

                let main_category = names.clone();
                println!("대분류: {}", main_category);

                // 상세 정보 수집 전 대기

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

                // 필수 데이터 검증
                if main_category.is_empty() || category.is_empty() || name.is_empty() {
                    println!(
                        "필수 데이터 누락: {} / {} / {}",
                        main_category, category, name
                    );
                    continue;
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
                                            features = feature_text
                                                .split('\n')
                                                .map(|line| format!("※ {}", line.trim()))
                                                .collect::<Vec<String>>()
                                                .join("\n");
                                        }
                                    }
                                } else if text.contains("제품 사용장소") {
                                    if let Ok(dd) = dl.find_element(By::Tag("dd")).await {
                                        if let Ok(usage_text) = dd.text().await {
                                            usage = usage_text
                                                .split('\n')
                                                .map(|line| format!("※ {}", line.trim()))
                                                .collect::<Vec<String>>()
                                                .join("\n");
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

                        match download_image(&full_url, &save_path).await {
                            Ok(_) => {
                                println!("이미지1 다운로드 성공: {}", save_path);
                                image1_path = save_path;
                            }
                            Err(e) => println!("이미지1 다운로드 실패: {:?}", e),
                        }
                    }
                }

                // 두 번째 이미지
                if let Ok(detail_img) = driver.find_element(By::Css(".detail .txt_area img")).await
                {
                    if let Ok(Some(src)) = detail_img.get_attribute("src").await {
                        let full_url = format!("{}{}", domain, src);
                        let file_name = src.split('/').last().unwrap_or("image2.jpg");
                        let save_path = format!("{}/2_{}", img_folder, file_name);

                        match download_image(&full_url, &save_path).await {
                            Ok(_) => {
                                println!("이미지2 다운로드 성공: {}", save_path);
                                image2_path = save_path;
                            }
                            Err(e) => println!("이미지2 다운로드 실패: {:?}", e),
                        }
                    }
                }

                // 엑셀에 데이터 저장
                sheet.write_string(row, 0, &names, None)?; // 대분류
                sheet.write_string(
                    row,
                    1,
                    if has_sub_categories {
                        &sub_name
                    } else {
                        &names
                    },
                    None,
                )?; // 세부분류
                sheet.write_string(row, 2, &name, None)?;
                sheet.write_string(row, 3, &features, None)?;
                sheet.write_string(row, 4, &usage, None)?;
                sheet.write_string(row, 5, &image1_path, None)?;
                sheet.write_string(row, 6, &image2_path, None)?;

                println!("데이터 저장 완료: row {}", row);
                row += 1;

                driver.goto(sub_href).await?;
                tokio::time::sleep(Duration::from_millis(200)).await;
            }
        }
    }

    println!("작업 완료. 엑셀 파일 저장 중...");
    workbook.close()?;

    driver.quit().await?;
    println!("크롤링 완료!");
    Ok(())
}

async fn download_image(url: &str, path: &str) -> Result<(), Box<dyn Error + Send + Sync>> {
    let response = reqwest::get(url).await?;
    let bytes = response.bytes().await?;
    fs::write(path, bytes)?;
    Ok(())
}
