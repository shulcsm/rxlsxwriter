use xml::writer::{EventWriter, XmlEvent};
use std::io::{Write};


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
