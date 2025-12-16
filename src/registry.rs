//! Windows Registry operations for COM object discovery.
//! 
//! This module handles all registry scanning and value retrieval operations,
//! including CLSID enumeration, ProgID lookup, and object descriptions.

use anyhow::Result;
use std::collections::HashMap;
use windows::core::{HSTRING, PCWSTR, PWSTR};
use windows::Win32::Foundation::ERROR_SUCCESS;
use windows::Win32::System::Registry::{
    RegCloseKey, RegEnumKeyExW, RegOpenKeyExW, RegQueryValueExW, HKEY, HKEY_CLASSES_ROOT,
    KEY_READ, REG_SAM_FLAGS, REG_VALUE_TYPE,
};

use crate::types::ComObject;
use crate::filter::should_include_object;

/// Scans the Windows registry for COM objects with specified filters
pub fn scan_com_objects(
    view_flag: REG_SAM_FLAGS,
    limit: usize,
    filter_description: &Option<String>,
    filter_clsid: &Option<String>,
    filter_app: &Option<Vec<String>>,
    interactive_filter: &Option<String>,
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

            // Check if this object passes all filters
            if should_include_object(
                &prog_id,
                &description,
                &clsid,
                interactive_filter,
                filter_description,
                filter_clsid,
                filter_app,
            ) {
                objects.insert(
                    clsid.clone(),
                    ComObject {
                        clsid,
                        prog_id,
                        description,
                    },
                );

                // Check limit
                if limit > 0 && objects.len() >= limit {
                    break;
                }
            }

            index += 1;
        }

        let _ = RegCloseKey(hkey_clsid);
    }

    Ok(objects)
}

/// Retrieves the ProgID for a given CLSID from the registry
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

/// Retrieves the description for a given CLSID from the registry
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

/// Low-level registry value reading with UTF-16 to UTF-8 conversion
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
