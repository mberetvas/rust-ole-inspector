use anyhow::Result;
use clap::Parser;
use std::collections::HashMap;
use windows::core::{HSTRING, PCWSTR, PWSTR};
use windows::Win32::Foundation::ERROR_SUCCESS;
use windows::Win32::System::Registry::{
    RegCloseKey, RegEnumKeyExW, RegOpenKeyExW, RegQueryValueExW, HKEY, HKEY_CLASSES_ROOT,
    KEY_READ, KEY_WOW64_32KEY, KEY_WOW64_64KEY, REG_SAM_FLAGS, REG_VALUE_TYPE,
};
// Console APIs used to enable UTF-8 output on Windows terminals
use windows::Win32::System::Console::{SetConsoleOutputCP, SetConsoleCP};

/// A Rust CLI for Windows that discovers COM objects and checks their programmatic usability
#[derive(Parser, Debug)]
#[command(author, version, about, long_about = None)]
struct Args {
    /// Show detailed information for each COM object
    #[arg(short, long)]
    verbose: bool,

    /// Scan 32-bit registry view
    #[arg(long)]
    scan_32bit: bool,

    /// Scan 64-bit registry view (default on 64-bit systems)
    #[arg(long)]
    scan_64bit: bool,

    /// Limit the number of results (0 = no limit)
    #[arg(short, long, default_value = "0")]
    limit: usize,

    /// Filter by ProgID, description, or CLSID substring (case-insensitive)
    #[arg(short, long)]
    filter: Option<String>,

    /// Filter by description substring only (case-insensitive)
    #[arg(long)]
    filter_description: Option<String>,

    /// Filter by CLSID substring only (case-insensitive)
    #[arg(long)]
    filter_clsid: Option<String>,
}

struct ComObject {
    clsid: String,
    prog_id: Option<String>,
    description: Option<String>,
}

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
        match scan_com_objects(view_flag, &args) {
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
    display_results(&all_objects, &args)?;

    Ok(())
}

fn check_privileges() {
    #[cfg(windows)]
    {
        use windows::Win32::Security::{GetTokenInformation, TokenElevation, TOKEN_ELEVATION};
        use windows::Win32::System::Threading::{GetCurrentProcess, OpenProcessToken};
        use windows::Win32::Foundation::HANDLE;
        use windows::Win32::Security::TOKEN_QUERY;

        unsafe {
            let mut token = HANDLE::default();
            if OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, &mut token).is_ok() {
                let mut elevation = TOKEN_ELEVATION::default();
                let mut return_length = 0u32;
                
                if GetTokenInformation(
                    token,
                    TokenElevation,
                    Some(&mut elevation as *mut _ as *mut _),
                    std::mem::size_of::<TOKEN_ELEVATION>() as u32,
                    &mut return_length,
                ).is_ok() {
                    if elevation.TokenIsElevated == 0 {
                        println!("⚠️  WARNING: Not running with elevated privileges.");
                        println!("   Some COM objects may not be accessible.");
                        println!("   Run as Administrator for a complete scan.\n");
                    } else {
                        println!("✓ Running with elevated privileges\n");
                    }
                }
            }
        }
    }
}

/// Try to enable UTF-8 output in the Windows console. Returns true on success.
fn init_console_utf8() -> bool {
    #[cfg(windows)]
    {
        unsafe {
            // 65001 = UTF-8 code page
            let out_ok = SetConsoleOutputCP(65001).is_ok();
            let in_ok = SetConsoleCP(65001).is_ok();
            out_ok && in_ok
        }
    }

    #[cfg(not(windows))]
    {
        // Non-Windows terminals usually handle UTF-8 fine
        true
    }
}

/// Print the original Unicode header art (uses box-drawing and other glyphs).
fn print_header_art_unicode() {
    let art = r#"
        ________              _____  ____________________________
        ___  __ \___  __________  /_ __  __ \__  /___  ____/__  /
        __  /_/ /  / / /_  ___/  __/ _  / / /_  / __  __/  __  / 
        _  _, _// /_/ /_(__  )/ /_   / /_/ /_  /___  /___   /_/  
        /_/ |_| \__,_/ /____/ \__/   \____/ /_____/_____/  (_)   
                                                                              
            Inspector - Discover and Analyze COM Objects

"#;

    println!("{art}");
}

/// Print a safe ASCII-only header as a fallback for terminals without UTF-8 support.
fn print_header_art_ascii() {
    let art = r#"
________              _____  ____________________________
___  __ \___  __________  /_ __  __ \__  /___  ____/__  /
__  /_/ /  / / /_  ___/  __/ _  / / /_  / __  __/  __  / 
_  _, _// /_/ /_(__  )/ /_   / /_/ /_  /___  /___   /_/  
/_/ |_| \__,_/ /____/ \__/   \____/ /_____/_____/  (_)       
            Rust COM Inspector - Discover and Analyze COM Objects

"#;

    println!("{art}");
}

fn scan_com_objects(
    view_flag: REG_SAM_FLAGS,
    args: &Args,
) -> Result<HashMap<String, ComObject>> {
    let mut objects = HashMap::new();

    unsafe {
        let clsid_path = HSTRING::from("CLSID");
        let mut hkey_clsid = HKEY::default();

        // Open HKEY_CLASSES_ROOT\CLSID with specified view
        let result = RegOpenKeyExW(
            HKEY_CLASSES_ROOT,
            &clsid_path,
            0,
            KEY_READ | view_flag,
            &mut hkey_clsid,
        );

        if result != ERROR_SUCCESS {
            return Err(anyhow::anyhow!(
                "Failed to open CLSID key: error code {}",
                result.0
            ));
        }

        // Enumerate all CLSIDs
        let mut index = 0u32;
        loop {
            let mut name_buffer = [0u16; 256];
            let mut name_len = name_buffer.len() as u32;

            let result = RegEnumKeyExW(
                hkey_clsid,
                index,
                PWSTR(name_buffer.as_mut_ptr()),
                &mut name_len,
                None,
                PWSTR::null(),
                None,
                None,
            );

            if result != ERROR_SUCCESS {
                break;
            }

            let clsid = String::from_utf16_lossy(&name_buffer[..name_len as usize]);

            // Try to get ProgID for this CLSID
            let prog_id = get_prog_id(hkey_clsid, &clsid);
            
            // Try to get description (default value)
            let description = get_description(hkey_clsid, &clsid);

            // Check if this object passes the filter
            if let Some(ref filter) = args.filter {
                let filter_lower = filter.to_lowercase();
                let matches = prog_id
                    .as_ref()
                    .map(|p| p.to_lowercase().contains(&filter_lower))
                    .unwrap_or(false)
                    || description
                        .as_ref()
                        .map(|d| d.to_lowercase().contains(&filter_lower))
                        .unwrap_or(false)
                    || clsid.to_lowercase().contains(&filter_lower);

                if !matches {
                    index += 1;
                    continue;
                }
            }

            // Check description filter
            if let Some(ref desc_filter) = args.filter_description {
                let desc_filter_lower = desc_filter.to_lowercase();
                let desc_matches = description
                    .as_ref()
                    .map(|d| d.to_lowercase().contains(&desc_filter_lower))
                    .unwrap_or(false);

                if !desc_matches {
                    index += 1;
                    continue;
                }
            }

            // Check CLSID filter
            if let Some(ref clsid_filter) = args.filter_clsid {
                let clsid_filter_lower = clsid_filter.to_lowercase();
                if !clsid.to_lowercase().contains(&clsid_filter_lower) {
                    index += 1;
                    continue;
                }
            }

            objects.insert(
                clsid.clone(),
                ComObject {
                    clsid,
                    prog_id,
                    description,
                },
            );

            // Check limit
            if args.limit > 0 && objects.len() >= args.limit {
                break;
            }

            index += 1;
        }

        let _ = RegCloseKey(hkey_clsid);
    }

    Ok(objects)
}

fn get_prog_id(hkey_clsid: HKEY, clsid: &str) -> Option<String> {
    unsafe {
        let progid_path = HSTRING::from(format!("{clsid}\\ProgID"));
        let mut hkey_progid = HKEY::default();

        if RegOpenKeyExW(hkey_clsid, &progid_path, 0, KEY_READ, &mut hkey_progid) == ERROR_SUCCESS
        {
            let value = read_registry_string(hkey_progid, None);
            let _ = RegCloseKey(hkey_progid);
            return value;
        }
    }
    None
}

fn get_description(hkey_clsid: HKEY, clsid: &str) -> Option<String> {
    unsafe {
        let clsid_path = HSTRING::from(clsid);
        let mut hkey_obj = HKEY::default();

        if RegOpenKeyExW(hkey_clsid, &clsid_path, 0, KEY_READ, &mut hkey_obj) == ERROR_SUCCESS {
            let value = read_registry_string(hkey_obj, None);
            let _ = RegCloseKey(hkey_obj);
            return value;
        }
    }
    None
}

fn read_registry_string(hkey: HKEY, value_name: Option<&str>) -> Option<String> {
    unsafe {
        let value_pcwstr = match value_name {
            Some(name) => {
                let hstring = HSTRING::from(name);
                PCWSTR(hstring.as_ptr())
            }
            None => PCWSTR::null(),
        };

        let mut size = 0u32;
        let mut type_code = REG_VALUE_TYPE(0);

        // Get size
        let result = RegQueryValueExW(
            hkey,
            value_pcwstr,
            None,
            Some(&mut type_code),
            None,
            Some(&mut size),
        );

        if result != ERROR_SUCCESS || size == 0 {
            return None;
        }

        // Read value
        let mut buffer = vec![0u16; (size / 2) as usize];
        let result = RegQueryValueExW(
            hkey,
            value_pcwstr,
            None,
            Some(&mut type_code),
            Some(buffer.as_mut_ptr() as *mut u8),
            Some(&mut size),
        );

        if result != ERROR_SUCCESS {
            return None;
        }

        // Remove null terminator and convert
        let len = buffer.iter().position(|&c| c == 0).unwrap_or(buffer.len());
        Some(String::from_utf16_lossy(&buffer[..len]))
    }
}

fn display_results(objects: &HashMap<String, ComObject>, args: &Args) -> Result<()> {
    println!("=== Results ===");
    println!("Total unique COM objects found: {}\n", objects.len());

    if objects.is_empty() {
        println!("No COM objects found matching the criteria.");
        return Ok(());
    }

    // Sort by ProgID for better readability
    let mut sorted_objects: Vec<_> = objects.values().collect();
    sorted_objects.sort_by(|a, b| {
        match (&a.prog_id, &b.prog_id) {
            (Some(pa), Some(pb)) => pa.cmp(pb),
            (Some(_), None) => std::cmp::Ordering::Less,
            (None, Some(_)) => std::cmp::Ordering::Greater,
            (None, None) => a.clsid.cmp(&b.clsid),
        }
    });

    // Count objects with ProgIDs
    let with_progid = sorted_objects
        .iter()
        .filter(|obj| obj.prog_id.is_some())
        .count();

    println!(
        "COM objects with ProgID: {} ({:.1}%)",
        with_progid,
        (with_progid as f64 / objects.len() as f64) * 100.0
    );
    println!("COM objects without ProgID: {}\n", objects.len() - with_progid);

    if args.verbose {
        println!("--- Detailed Listing ---\n");
        for obj in sorted_objects {
            println!("CLSID: {}", obj.clsid);
            if let Some(ref prog_id) = obj.prog_id {
                println!("  ProgID: {prog_id}");
            }
            if let Some(ref desc) = obj.description {
                println!("  Description: {desc}");
            }
            
            // Check programmatic usability
            let usability = check_usability(obj);
            println!("  Programmatic Usability: {usability}");
            println!();
        }
    } else {
        // Just show ProgIDs in compact format
        println!("--- COM Objects with ProgID ---");
        for obj in sorted_objects.iter().filter(|o| o.prog_id.is_some()) {
            if let Some(ref prog_id) = obj.prog_id {
                println!("  {} ({})", prog_id, obj.clsid);
            }
        }
    }

    Ok(())
}

fn check_usability(obj: &ComObject) -> &'static str {
    // An object is more likely to be programmatically usable if:
    // 1. It has a ProgID (can be instantiated by name)
    // 2. It has a description (indicates it's documented)
    match (&obj.prog_id, &obj.description) {
        (Some(_), Some(_)) => "✓ High (has ProgID and description)",
        (Some(_), None) => "~ Medium (has ProgID)",
        (None, Some(_)) => "~ Low (no ProgID, has description)",
        (None, None) => "✗ Very Low (no ProgID or description)",
    }
}
