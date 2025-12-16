mod console;
mod display;
mod filter;
mod registry;
mod security;
mod types;

use anyhow::Result;
use clap::Parser;
use std::collections::HashMap;
use std::io;
use windows::Win32::System::Registry::{KEY_WOW64_32KEY, KEY_WOW64_64KEY};

use console::{init_console_utf8, print_header_art_ascii, print_header_art_unicode};
use display::{display_results, prompt_export};
use registry::scan_com_objects;
use security::check_privileges;
use types::{Args, ComObject};

fn main() -> Result<()> {
    let args = Args::parse();

    // Check if running with elevated privileges
    check_privileges();

    // Try to enable UTF-8 console output; fall back to ASCII art if unavailable
    let unicode_ok = init_console_utf8();
    if unicode_ok {
        print_header_art_unicode();
    } else {
        print_header_art_ascii();
    }

    // Prompt user for filter
    println!("Enter a filter for COM objects (leave empty to search all):");
    let mut interactive_filter = String::new();
    io::stdin().read_line(&mut interactive_filter)?;
    let interactive_filter = interactive_filter.trim().to_string();
    let interactive_filter = if interactive_filter.is_empty() {
        None
    } else {
        Some(interactive_filter)
    };

    // Determine which registry views to scan
    let mut views_to_scan = Vec::new();

    if args.scan_32bit || !args.scan_64bit {
        views_to_scan.push(("32-bit", KEY_WOW64_32KEY));
    }

    if args.scan_64bit || !args.scan_32bit {
        views_to_scan.push(("64-bit", KEY_WOW64_64KEY));
    }

    let mut all_objects = HashMap::new();

    for (view_name, view_flag) in views_to_scan {
        println!("Scanning {view_name} registry view...");
        match scan_com_objects(
            view_flag,
            args.limit,
            &args.filter_description,
            &args.filter_clsid,
            &args.filter_app,
            &interactive_filter,
        ) {
            Ok(objects) => {
                println!("Found {} COM objects in {} view\n", objects.len(), view_name);

                // Merge objects, preferring those with ProgIDs
                for (clsid, obj) in objects {
                    all_objects
                        .entry(clsid.clone())
                        .and_modify(|existing: &mut ComObject| {
                            if obj.prog_id.is_some() && existing.prog_id.is_none() {
                                existing.prog_id = obj.prog_id.clone();
                            }
                            if obj.description.is_some() && existing.description.is_none() {
                                existing.description = obj.description.clone();
                            }
                        })
                        .or_insert(obj);
                }
            }
            Err(e) => {
                eprintln!("Error scanning {view_name} view: {e}");
            }
        }
    }

    // Display results
    display_results(&all_objects, args.verbose)?;

    prompt_export(&all_objects)?;

    // Wait for user to press 'q' to quit
    println!("Press 'q' to quit...");
    let stdin = io::stdin();
    for line in stdin.lines().map_while(Result::ok) {
        if line.trim().to_lowercase() == "q" {
            break;
        }
    }

    Ok(())
}
