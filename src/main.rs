use eframe::egui;

fn main() -> eframe::Result<()> {
    let app = MyApp::default();

    let options = eframe::NativeOptions::default();
    eframe::run_native("jang sung jin", options, Box::new(|_cc: &eframe::CreationContext<'_>| Ok(Box::new(app))))
}

struct MyApp {
    shown: bool,
}

impl Default for MyApp {
    fn default() -> Self {
        Self { shown: false }
    }
}

impl eframe::App for MyApp {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        egui::CentralPanel::default().show(ctx, |ui| {
            if ui.button("Click Me").clicked() {
                self.shown = true;
            }
            if self.shown {
                ui.label("Hello World!");
            }
        });
    }
}