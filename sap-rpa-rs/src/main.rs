
// use winput;
use std::io;
use std::io::{BufReader, prelude::*};
use std::fs::File;

fn main() -> io::Result<()> {

    let f = File::open("orders.txt")?;
    let reader = BufReader::new(f);

    // read a line into buffer
    for line in reader.lines() {
        println!("Order: {}", line?);
    }

    Ok(())
}
