//! Display and export functionality for COM object results.
//! 
//! This module handles result presentation to the user and exporting to various formats (txt, csv).

use anyhow::Result;
use std::collections::HashMap;
use std::fs::File;
use std::io::Write;
use csv::Writer;

use crate::types::ComObject;

/// Display results to console with optional verbose output
pub fn display_results(objects: &HashMap<String, ComObject>, verbose: bool) -> Result<()> {
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

    if verbose {
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

/// Prompt user for export options and perform export
pub fn prompt_export(objects: &HashMap<String, ComObject>) -> Result<()> {
    println!("Do you want to export the results? (y/n): ");
    let mut input = String::new();
    std::io::stdin().read_line(&mut input)?;
    if input.trim().to_lowercase() != "y" {
        return Ok(());
    }

    println!("Export format (txt/csv): ");
    let mut format_input = String::new();
    std::io::stdin().read_line(&mut format_input)?;
    let format = format_input.trim().to_lowercase();
    if format != "txt" && format != "csv" {
        println!("Invalid format, skipping export.");
        return Ok(());
    }

    println!("Enter file path to export to: ");
    let mut path_input = String::new();
    std::io::stdin().read_line(&mut path_input)?;
    let path = path_input.trim();

    if path.is_empty() {
        println!("Export cancelled: No file path provided.");
        return Ok(());
    }

    // Use a match block to handle errors instead of '?'
    // This prevents the program from exiting immediately on "Access Denied" errors
    let export_result = if format == "txt" {
        export_txt(objects, path)
    } else {
        export_csv(objects, path)
    };

    match export_result {
        Ok(_) => println!("Successfully exported to {}", path),
        Err(e) => {
            println!("\n❌ Export failed: {}", e);
            println!("Hint: If running as Administrator, writing to the default folder (System32) is restricted.");
            println!("      Try providing a full absolute path (e.g., C:\\Temp\\results.{})", format);
        }
    }

    Ok(())
}

/// Export results to a text file
fn export_txt(objects: &HashMap<String, ComObject>, path: &str) -> Result<()> {
    let mut output = String::new();
    output.push_str("=== Results ===\n");
    output.push_str(&format!("Total unique COM objects found: {}\n\n", objects.len()));

    if objects.is_empty() {
        output.push_str("No COM objects found matching the criteria.\n");
    } else {
        // Sort by ProgID
        let mut sorted_objects: Vec<_> = objects.values().collect();
        sorted_objects.sort_by(|a, b| {
            match (&a.prog_id, &b.prog_id) {
                (Some(pa), Some(pb)) => pa.cmp(pb),
                (Some(_), None) => std::cmp::Ordering::Less,
                (None, Some(_)) => std::cmp::Ordering::Greater,
                (None, None) => a.clsid.cmp(&b.clsid),
            }
        });

        let with_progid = sorted_objects
            .iter()
            .filter(|obj| obj.prog_id.is_some())
            .count();
        output.push_str(&format!(
            "COM objects with ProgID: {} ({:.1}%)\n",
            with_progid,
            (with_progid as f64 / objects.len() as f64) * 100.0
        ));
        output.push_str(&format!(
            "COM objects without ProgID: {}\n\n",
            objects.len() - with_progid
        ));
        output.push_str("--- Detailed Listing ---\n\n");

        for obj in sorted_objects {
            output.push_str(&format!("CLSID: {}\n", obj.clsid));
            if let Some(ref prog_id) = obj.prog_id {
                output.push_str(&format!("  ProgID: {}\n", prog_id));
            }
            if let Some(ref desc) = obj.description {
                output.push_str(&format!("  Description: {}\n", desc));
            }
            let usability = check_usability(obj);
            output.push_str(&format!("  Programmatic Usability: {}\n\n", usability));
        }
    }

    let mut file = File::create(path)?;
    file.write_all(output.as_bytes())?;
    Ok(())
}

/// Export results to a CSV file
fn export_csv(objects: &HashMap<String, ComObject>, path: &str) -> Result<()> {
    let mut wtr = Writer::from_writer(File::create(path)?);
    wtr.write_record(&["CLSID", "ProgID", "Description", "Usability"])?;

    let mut sorted_objects: Vec<_> = objects.values().collect();
    sorted_objects.sort_by(|a, b| {
        match (&a.prog_id, &b.prog_id) {
            (Some(pa), Some(pb)) => pa.cmp(pb),
            (Some(_), None) => std::cmp::Ordering::Less,
            (None, Some(_)) => std::cmp::Ordering::Greater,
            (None, None) => a.clsid.cmp(&b.clsid),
        }
    });

    for obj in sorted_objects {
        let usability = check_usability(obj);
        wtr.write_record(&[
            obj.clsid.as_str(),
            obj.prog_id.as_deref().unwrap_or(""),
            obj.description.as_deref().unwrap_or(""),
            usability,
        ])?;
    }

    wtr.flush()?;
    Ok(())
}

/// Assess programmatic usability of a COM object
pub fn check_usability(obj: &ComObject) -> &'static str {
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
