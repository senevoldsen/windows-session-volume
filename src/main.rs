pub mod lib;

use volume_fix::{find_application_for_device, find_device_with_friendly_name, set_volume};
use windows::Win32::System::Com::{CoInitializeEx, COINIT_MULTITHREADED};


fn change_volume() -> Result<(), &'static str> {
    unsafe {
        CoInitializeEx(std::ptr::null_mut(), COINIT_MULTITHREADED).unwrap();
    }

    let device_name = std::env::args().nth(1).ok_or("Missing device name")?;
    let application_name = std::env::args()
        .nth(2)
        .ok_or("Missing application/session name")?;
    let volume_factor: f32 = std::env::args()
        .nth(3)
        .ok_or("Missing volume to set to")?
        .parse()
        .map_err(|_| "Invalid number for volume")?;

    if volume_factor < 0f32 || volume_factor > 1f32 {
        return Err("Volume must be between 0 and 1");
    }

    let mut found_device = false;
    let device = find_device_with_friendly_name(&device_name);
    match &device {
        Ok(d) => {
            println!("Found Device");
            let session = find_application_for_device(d, |app_name| app_name == &application_name);
            match &session {
                Ok(s) => {
                    found_device = true;
                    println!("Setting volume");
                    set_volume(s, volume_factor);
                }
                Err(e) => {
                    println!("Failed to find audio session");
                    println!("Error: {}", e);
                }
            }
        }
        Err(e) => println!("Error: {}", e),
    };

    if found_device {
        Ok(())
    } else {
        Err("Device not found")
    }
}

fn main() {
    let result = change_volume();
    match result {
        Ok(_) => {
            std::process::exit(0);
        }
        Err(s) => {
            println!("Error {}", s);
            std::process::exit(1);
        }
    }
}
