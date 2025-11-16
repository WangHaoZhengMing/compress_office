/// 带图片压缩率的 PPTX 压缩
pub fn compress_pptx_with_quality<F>(
    input_path: &str, 
    output_path: &str, 
    image_quality: f32,
    progress_callback: F
) -> Result<String> 
where
    F: Fn(usize, usize) + Send + 'static,
{
    let start_time = std::time::Instant::now();
    
    let input_file = File::open(input_path)
        .context("无法打开输入文件")?;
    let mut archive = ZipArchive::new(input_file)
        .context("无法解析 PPTX 文件（可能不是有效的 PPTX 格式）")?;
    
    // 先统计总图片数（收集为拥有的文件名，避免引用逃逸）
    let total_images = (0..archive.len())
        .filter_map(|i| archive.by_index(i).ok().map(|f| f.name().to_string()))
        .filter(|name| is_image_file(name))
        .count();
    
    let output_file = File::create(output_path)
        .context("无法创建输出文件")?;
    let mut zip_writer = ZipWriter::new(output_file);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(CompressionMethod::Deflated)
        .compression_level(Some(9));
    
    let mut stats = CompressionStats::default();
    stats.total_files = archive.len();
    let mut processed_images = 0;
    
    for i in 0..archive.len() {
        let mut file = archive.by_index(i)?;
        let name = file.name().to_owned();
        zip_writer.start_file(&name, options)?;
        if name.ends_with(".xml") || name.ends_with(".rels") {
            let mut contents = String::new();
            file.read_to_string(&mut contents)?;
            let original_len = contents.len();
            let optimized = optimize_xml(&contents);
            let saved = original_len.saturating_sub(optimized.len());
            stats.xml_files += 1;
            stats.xml_saved += saved;
            zip_writer.write_all(optimized.as_bytes())?;
        } else if is_image_file(&name) {
            processed_images += 1;
            progress_callback(processed_images, total_images);
            
            let mut buffer = Vec::new();
            file.read_to_end(&mut buffer)?;
            let original_len = buffer.len();
            match compress_image(&buffer, image_quality) {
                Ok(img) => {
                    let saved = original_len.saturating_sub(img.len());
                    stats.images_compressed += 1;
                    stats.image_saved += saved;
                    zip_writer.write_all(&img)?;
                }
                Err(_) => {
                    stats.images_skipped += 1;
                    zip_writer.write_all(&buffer)?;
                }
            }
        } else {
            let mut buffer = Vec::new();
            file.read_to_end(&mut buffer)?;
            zip_writer.write_all(&buffer)?;
        }
    }
    zip_writer.finish()?;
    
    let elapsed = start_time.elapsed();
    let original_size = std::fs::metadata(input_path)?.len();
    let compressed_size = std::fs::metadata(output_path)?.len();
    let saved = original_size.saturating_sub(compressed_size);
    let percent = if original_size > 0 {
        (saved as f64 / original_size as f64 * 100.0) as i32
    } else {
        0
    };
    
    Ok(format!(
        "✓ 压缩完成！\n\n\
        📊 文件信息:\n\
        • 原始大小: {:.2} MB ({} KB)\n\
        • 压缩后: {:.2} MB ({} KB)\n\
        • 节省空间: {:.2} MB ({} KB)\n\
        • 压缩率: {}%\n\n\
        📁 处理统计:\n\
        • 总文件数: {}\n\
        • XML文件: {} 个 (节省 {:.1} KB)\n\
        • 图片压缩: {} 个 (节省 {:.1} KB)\n\
        • 图片跳过: {} 个\n\
        • 图片质量: {}%\n\n\
        ⏱️ 处理耗时: {:.2} 秒",
        original_size as f64 / 1024.0 / 1024.0,
        original_size / 1024,
        compressed_size as f64 / 1024.0 / 1024.0,
        compressed_size / 1024,
        saved as f64 / 1024.0 / 1024.0,
        saved / 1024,
        percent,
        stats.total_files,
        stats.xml_files,
        stats.xml_saved as f64 / 1024.0,
        stats.images_compressed,
        stats.image_saved as f64 / 1024.0,
        stats.images_skipped,
        (image_quality * 100.0) as u8,
        elapsed.as_secs_f64()
    ))
}

#[derive(Default)]
struct CompressionStats {
    total_files: usize,
    xml_files: usize,
    xml_saved: usize,
    images_compressed: usize,
    images_skipped: usize,
    image_saved: usize,
}
use anyhow::{Context, Result};
use std::fs::File;
use std::io::{Read, Write};
use zip::{ZipArchive, ZipWriter, CompressionMethod};

/// 压缩 PPTX 文件
/// 
/// 原理：PPTX 也是 ZIP 格式，包含 XML 文件、图片、主题等资源
/// 压缩策略：
/// 1. 使用最大压缩级别重新打包
/// 2. 优化 XML 文件
/// 3. 压缩图片资源（未来可以添加图片质量压缩）
#[warn(dead_code)]
pub fn compress_pptx(input_path: &str, output_path: &str) -> Result<String> {
    let input_file = File::open(input_path)
        .context("无法打开输入文件")?;
    let mut archive = ZipArchive::new(input_file)
        .context("无法解析 PPTX 文件（可能不是有效的 PPTX 格式）")?;
    let output_file = File::create(output_path)
        .context("无法创建输出文件")?;
    let mut zip_writer = ZipWriter::new(output_file);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(CompressionMethod::Deflated)
        .compression_level(Some(9)); // 最大压缩级别

    // 用户可调节图片压缩率，范围 0.0~1.0，默认 0.8
    let image_quality = 0.8; // TODO: 从 UI 传入

    for i in 0..archive.len() {
        let mut file = archive.by_index(i)?;
        let name = file.name().to_owned();
        zip_writer.start_file(&name, options)?;
        if name.ends_with(".xml") || name.ends_with(".rels") {
            let mut contents = String::new();
            file.read_to_string(&mut contents)?;
            let optimized = optimize_xml(&contents);
            zip_writer.write_all(optimized.as_bytes())?;
        } else if is_image_file(&name) {
            let mut buffer = Vec::new();
            file.read_to_end(&mut buffer)?;
            // 压缩图片
            match compress_image(&buffer, image_quality) {
                Ok(img) => zip_writer.write_all(&img)?,
                Err(_) => zip_writer.write_all(&buffer)?, // 压缩失败则原样写入
            }
        } else {
            let mut buffer = Vec::new();
            file.read_to_end(&mut buffer)?;
            zip_writer.write_all(&buffer)?;
        }
    }
    zip_writer.finish()?;
    let original_size = std::fs::metadata(input_path)?.len();
    let compressed_size = std::fs::metadata(output_path)?.len();
    let saved = original_size.saturating_sub(compressed_size);
    let percent = if original_size > 0 {
        (saved as f64 / original_size as f64 * 100.0) as i32
    } else {
        0
    };
    Ok(format!(
        "压缩完成！\n原始大小: {} KB\n压缩后: {} KB\n节省: {} KB ({} %)",
        original_size / 1024,
        compressed_size / 1024,
        saved / 1024,
        percent
    ))
}
fn compress_image(data: &[u8], quality: f32) -> Result<Vec<u8>> {
    use image::ImageReader;
    use image::ImageEncoder;
    use image::codecs::jpeg::JpegEncoder;
    use image::codecs::png::{PngEncoder, CompressionType, FilterType};
    use std::io::Cursor;
    
    // 检测原始格式
    let format = ImageReader::new(Cursor::new(data))
        .with_guessed_format()
        .map_err(|_| anyhow::anyhow!("图片格式检测失败"))?
        .format();
    
    let img = ImageReader::new(Cursor::new(data))
        .with_guessed_format()
        .map_err(|_| anyhow::anyhow!("图片解码失败"))?
        .decode()
        .map_err(|_| anyhow::anyhow!("图片解码失败"))?;
    
    let mut buf = Cursor::new(Vec::new());
    
    // 根据原始格式进行压缩，保持格式不变
    match format {
        Some(image::ImageFormat::Png) => {
            // PNG 格式：保留透明通道，使用适当压缩
            let encoder = PngEncoder::new_with_quality(
                &mut buf,
                CompressionType::Best,
                FilterType::Adaptive,
            );
            encoder.write_image(
                img.as_bytes(),
                img.width(),
                img.height(),
                img.color().into(),
            ).map_err(|_| anyhow::anyhow!("PNG 编码失败"))?;
        }
        Some(image::ImageFormat::Jpeg) => {
            // JPEG 格式：按质量压缩
            let quality_u8 = (quality * 100.0).round() as u8;
            let mut encoder = JpegEncoder::new_with_quality(&mut buf, quality_u8);
            encoder.encode_image(&img)
                .map_err(|_| anyhow::anyhow!("JPEG 编码失败"))?;
        }
        _ => {
            // 其他格式：不压缩，返回原始数据
            return Err(anyhow::anyhow!("不支持的图片格式，保持原样"));
        }
    }
    
    let compressed = buf.into_inner();
    
    // 如果压缩后更大，则使用原始数据
    if compressed.len() >= data.len() {
        return Err(anyhow::anyhow!("压缩后不减小，保持原样"));
    }
    
    Ok(compressed)
}

/// 判断是否为图片文件
fn is_image_file(filename: &str) -> bool {
    let lower = filename.to_lowercase();
    lower.ends_with(".png") 
        || lower.ends_with(".jpg") 
        || lower.ends_with(".jpeg") 
        || lower.ends_with(".gif")
        || lower.ends_with(".bmp")
        || lower.ends_with(".emf")
        || lower.ends_with(".wmf")
}

/// 优化 XML 内容
/// 移除多余的空白符和换行，但保留必要的格式
fn optimize_xml(xml: &str) -> String {
    xml.lines()
        .map(|line| line.trim())
        .filter(|line| !line.is_empty())
        .collect::<Vec<_>>()
        .join("")
}

#[cfg(test)]
mod tests {
    use super::*;
    
    #[test]
    fn test_is_image_file() {
        assert!(is_image_file("slide1/media/image1.png"));
        assert!(is_image_file("ppt/media/image2.JPG"));
        assert!(!is_image_file("slide1.xml"));
    }
    
    #[test]
    fn test_xml_optimization() {
        let input = r#"
        <presentation>
            <slide>
                <content>Test</content>
            </slide>
        </presentation>
        "#;
        
        let output = optimize_xml(input);
        assert!(!output.contains('\n'));
        assert!(output.contains("<presentation>"));
    }
}
