use eframe::egui;
use std::error::Error;
use std::thread;
use std::time::Duration;
use thirtyfour::prelude::*;
use tokio::process::Command;
use tokio::time::sleep;
use dotenv::dotenv;
use std::env;


const WINDOW_DRIVER: &str = "./driver/chromedriver.exe";
const MAC_DRIVER: &str = "./driver/chromedriver";

const TYPE: &str = "WINDOW";
fn main() -> eframe::Result<()> {
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
    Command::new(chromedriver_path).arg("--port=9515").spawn()?;

    sleep(Duration::from_secs(5)).await;
    Ok(())
}

pub async fn example() -> Result<(), Box<dyn Error + Send + Sync>> {
    start_chromedriver().await?;

    let mut caps = DesiredCapabilities::chrome();
    caps.set_no_sandbox()?;

    let driver = WebDriver::new("http://localhost:9515", caps).await?;
    driver.goto("https://ikor250113s.mycafe24.com/").await?;
    let admin_id = env::var("ADMIN_ID").expect("WINDOW_DRIVER must be set");
    let admin_pw = env::var("ADMIN_PW").expect("WINDOW_DRIVER must be set");

    println!("{}",admin_id);
    println!("{}",admin_pw);

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

            if ui.button("Button 2").clicked() {
                self.shown2 = !self.shown2;
            }
            if self.shown2 {
                ui.label("Hello from Button 2!");
            }
        });
    }
}
