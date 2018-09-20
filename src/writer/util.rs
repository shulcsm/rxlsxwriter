use std::io::{Write, Seek};
use xml::writer::{EventWriter, XmlEvent};
use xml::{EmitterConfig};
use zip::{ZipWriter, CompressionMethod, write::FileOptions};


// @TODO try converting this to extension trait
pub fn xml_file<'a>(zip_writer: &'a mut ZipWriter<impl Write + Seek>, file_name: &str) -> EventWriter<impl Write + 'a> {
    let options = FileOptions::default().compression_method(CompressionMethod::Deflated);
    zip_writer.start_file(file_name, options).unwrap();
    EmitterConfig::new().perform_indent(true)
        .create_writer(zip_writer)
}

pub trait XmlWriter {
    fn write_and_close<'a, E>(&mut self, event: E) -> () where E: Into<XmlEvent<'a>>;
    fn write_and_close_chars<'a, E>(&mut self, event: E, chars: &str) -> () where E: Into<XmlEvent<'a>>;
}

impl<W: Write> XmlWriter for EventWriter<W> {
    fn write_and_close<'a, E>(&mut self, event: E) -> () where E: Into<XmlEvent<'a>> {
        self.write(event);
        self.write(XmlEvent::end_element());
    }

    fn write_and_close_chars<'a, E>(&mut self, event: E, chars: &str) -> () where E: Into<XmlEvent<'a>> {
        self.write(event).unwrap();
        self.write(XmlEvent::Characters(chars));
        self.write(XmlEvent::end_element());
    }
}
