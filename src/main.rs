use std::process::ExitCode;

use volume_fix::{find_application_for_device, find_device_with_friendly_name, set_volume};
use windows::Win32::System::Com::{CoInitializeEx, COINIT_MULTITHREADED};

fn change_volume(args: &[&str]) -> Result<(), &'static str> {
    unsafe {
        CoInitializeEx(None, COINIT_MULTITHREADED).unwrap();
    }

    let device_name = *args.get(0).ok_or("Missing device name")?;
    let application_name = *args.get(1).ok_or("Missing application/session name")?;
    let volume_factor: f32 = args
        .get(2)
        .ok_or("Missing volume to set to")?
        .parse()
        .map_err(|_| "Invalid number for volume")?;

    if !(0f32..=1f32).contains(&volume_factor) {
        return Err("Volume must be between 0 and 1");
    }

    let mut found_device = false;
    let device = find_device_with_friendly_name(&device_name);
    match &device {
        Ok(d) => {
            println!("Found Device");
            let session = find_application_for_device(d, |app_name| app_name == application_name);
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

fn print_usage(progname: &str) {
    println!(
        "USAGE: #{progname} <COMMAND> [args...]

COMMANDS:

help/--help
    This message

set-volume <speaker-device-name> <app/session-name> <volume>
    Sets the volume for the session with the device to a volume between 0.0 and 1.0
"
    )
}

fn main() -> ExitCode {
    let args: Vec<String> = std::env::args().collect();
    let args: Vec<&str> = args.iter().map(AsRef::as_ref).collect();

    if args.len() < 2 {
        print_usage(&args[0]);
        return ExitCode::from(1);
    }

    let progname = &args[0];
    let command = &args[1];
    let cmd_args = &args[2..];
    let result = match *command {
        "help" | "--help" => {
            print_usage(progname);
            Ok(())
        }
        "set-volume" => change_volume(cmd_args),
        _ => {
            print_usage(progname);
            return ExitCode::SUCCESS;
        }
    };
    if result.is_err() {
        let err = result.unwrap_err();
        println!("ERROR: #{err}");
        ExitCode::FAILURE
    } else {
        ExitCode::SUCCESS
    }
}
