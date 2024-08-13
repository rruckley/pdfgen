use docx::Docx;
use serde::Deserialize;
use serde_json::Value;
use reqwest::Error;
use std::fs::File;
use std::io::{Read, Write};
use printpdf::*;
use std::io::BufWriter;
use jsonpath_lib as jsonpath;
#[derive(Deserialize)]
struct Config {
   mappings: std::collections::HashMap<String, String>,
}
#[tokio::main]
async fn main() -> Result<(), Error> {
   // Read the configuration file
   let config_file = File::open("config.json").expect("Cannot open config file");
   let config: Config = serde_json::from_reader(config_file).expect("Cannot read config file");
   // Fetch the replacements from the web service
   let response = reqwest::get("https://example.com/replacements")
       .await?
       .json::<Value>()
       .await?;
   // Open the template document
   let mut doc = Docx::open("path/to/your/template.docx").expect("Cannot open template file");
   // Read the document content
   let mut content = String::new();
   doc.read_to_string(&mut content).expect("Cannot read template file");
   // Perform the replacements
   for (placeholder, json_path) in config.mappings {
       if let Ok(replacement) = jsonpath::select(&response, &json_path) {
           if let Some(replacement_value) = replacement.get(0) {
               content = content.replace(&format!("{{{{{}}}}}", placeholder), replacement_value.as_str().unwrap_or(""));
           }
       }
   }
   // Write the modified content back to a new file
   let mut output = File::create("path/to/your/output.docx").expect("Cannot create output file");
   output.write_all(content.as_bytes()).expect("Cannot write to output file");
   // Convert the content to a PDF
   let (doc, page1, layer1) = PdfDocument::new("PDF_Document_title", Mm(210.0), Mm(297.0), "Layer 1");
   let current_layer = doc.get_page(page1).get_layer(layer1);
   let font = doc.add_external_font(File::open("path/to/your/font.ttf").unwrap()).unwrap();
   let text = content;
   current_layer.use_text(text, 12.0, Mm(10.0), Mm(287.0), &font);
   doc.save(&mut BufWriter::new(File::create("output.pdf").unwrap())).unwrap();
   println!("Mail merge and PDF conversion completed successfully!");
   Ok(())
}