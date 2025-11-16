// 在 Windows Release 模式下隐藏控制台窗口
#![cfg_attr(all(windows, not(debug_assertions)), windows_subsystem = "windows")]


mod docx;
mod pptx;

use std::path::Path;

slint::include_modules!();

fn main() {
    let ui = MainWindow::new().unwrap();
    
    // 克隆 UI 引用用于回调
    let ui_weak = ui.as_weak();
    
    // 文件选择回调
    ui.on_select_file(move || {
        let ui = ui_weak.unwrap();
        
        // 使用 rfd (Rust File Dialog) 创建文件选择对话框
        // 注意：需要在 Cargo.toml 中添加 rfd 依赖
        if let Some(path) = native_dialog::FileDialog::new()
            .add_filter("Office 文档", &["docx", "pptx"])
            .show_open_single_file()
            .ok()
            .flatten()
        {
            ui.set_file_path(path.to_string_lossy().to_string().into());
            ui.set_status_text("已选择文件，点击「开始压缩」按钮".into());
            ui.set_progress(0.0);
        }
    });
    
    // 克隆另一个 UI 引用用于压缩回调
    let ui_weak = ui.as_weak();
    
    // 压缩文件回调
    ui.on_compress_file(move || {
        let ui = ui_weak.unwrap();
        let input_path = ui.get_file_path().to_string();
        
        if input_path.is_empty() {
            ui.set_status_text("请先选择一个文件！".into());
            return;
        }
        
        // 设置处理状态
        ui.set_is_processing(true);
        ui.set_current_step("📂 正在读取文件...".into());
        ui.set_status_text("开始处理，请稍候...".into());
        ui.set_progress(0.1);
        
        // 生成输出文件名
        let path = Path::new(&input_path);
        let output_path = if let Some(stem) = path.file_stem() {
            let parent = path.parent().unwrap_or(Path::new("."));
            let ext = path.extension().and_then(|s| s.to_str()).unwrap_or("");
            parent.join(format!("{}_compressed.{}", stem.to_string_lossy(), ext))
        } else {
            path.with_extension("compressed")
        };
        
        // 获取图片压缩率
        let image_quality = ui.get_image_quality();
        
        // 克隆 UI 引用用于后台线程
        let ui_handle = ui.as_weak();
        let output_path_clone = output_path.clone();
        
        // 在后台线程执行压缩任务，避免阻塞 UI
        std::thread::spawn(move || {
            // 步骤1: 解析文件
            let ui = ui_handle.clone();
            slint::invoke_from_event_loop(move || {
                let ui = ui.unwrap();
                ui.set_current_step("🔍 解析文档结构...".into());
                ui.set_progress(0.2);
            }).ok();
            std::thread::sleep(std::time::Duration::from_millis(300));
            
            // 步骤2: 优化XML
            let ui = ui_handle.clone();
            slint::invoke_from_event_loop(move || {
                let ui = ui.unwrap();
                ui.set_current_step("📝 优化 XML 文件...".into());
                ui.set_progress(0.35);
            }).ok();
            std::thread::sleep(std::time::Duration::from_millis(200));
            
            // 步骤3: 压缩图片
            let ui = ui_handle.clone();
            slint::invoke_from_event_loop(move || {
                let ui = ui.unwrap();
                ui.set_current_step("🖼️ 压缩图片资源...".into());
                ui.set_progress(0.5);
            }).ok();
            
            // 执行压缩（带进度回调）
            let ui_progress = ui_handle.clone();
            let result = if input_path.to_lowercase().ends_with(".docx") {
                docx::compress_docx_with_quality(
                    &input_path, 
                    output_path_clone.to_str().unwrap(), 
                    image_quality,
                    move |processed, total| {
                        let ui = ui_progress.clone();
                        slint::invoke_from_event_loop(move || {
                            let ui = ui.unwrap();
                            let remaining = total.saturating_sub(processed);
                            ui.set_current_step(format!("🖼️ 压缩图片... ({}/{}，剩余 {})", processed, total, remaining).into());
                            ui.set_total_images(total as i32);
                            ui.set_processed_images(processed as i32);
                        }).ok();
                    }
                )
            } else if input_path.to_lowercase().ends_with(".pptx") {
                pptx::compress_pptx_with_quality(
                    &input_path, 
                    output_path_clone.to_str().unwrap(), 
                    image_quality,
                    move |processed, total| {
                        let ui = ui_progress.clone();
                        slint::invoke_from_event_loop(move || {
                            let ui = ui.unwrap();
                            let remaining = total.saturating_sub(processed);
                            ui.set_current_step(format!("🖼️ 压缩图片... ({}/{}，剩余 {})", processed, total, remaining).into());
                            ui.set_total_images(total as i32);
                            ui.set_processed_images(processed as i32);
                        }).ok();
                    }
                )
            } else {
                Err(anyhow::anyhow!("不支持的文件格式，仅支持 .docx 和 .pptx"))
            };
            
            // 步骤4: 重新打包
            let ui = ui_handle.clone();
            slint::invoke_from_event_loop(move || {
                let ui = ui.unwrap();
                ui.set_current_step("📦 重新打包文件...".into());
                ui.set_progress(0.85);
            }).ok();
            std::thread::sleep(std::time::Duration::from_millis(200));
            
            // 步骤5: 完成
            slint::invoke_from_event_loop(move || {
                let ui = ui_handle.unwrap();
                ui.set_current_step("✅ 处理完成！".into());
                ui.set_progress(1.0);
                
                // 显示结果
                match result {
                    Ok(msg) => {
                        let full_msg = format!(
                            "{}\n\n输出文件: {}",
                            msg,
                            output_path_clone.display()
                        );
                        ui.set_status_text(full_msg.into());
                    }
                    Err(e) => {
                        ui.set_status_text(format!("压缩失败: {}", e).into());
                        ui.set_current_step("❌ 处理失败".into());
                        ui.set_progress(0.0);
                    }
                }
                
                // 延迟清除步骤提示
                let ui_weak = ui.as_weak();
                std::thread::spawn(move || {
                    std::thread::sleep(std::time::Duration::from_secs(2));
                    slint::invoke_from_event_loop(move || {
                        let ui = ui_weak.unwrap();
                        ui.set_is_processing(false);
                        ui.set_current_step("".into());
                    }).ok();
                });
            }).ok();
        });
    });
    
    ui.run().unwrap();
}

