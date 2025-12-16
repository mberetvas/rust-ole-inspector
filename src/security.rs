//! Windows security utilities for privilege checking.
//! 
//! This module handles detection of elevated privileges and warnings if unavailable.

/// Check if running with elevated privileges and print appropriate warning/confirmation.
pub fn check_privileges() {
    #[cfg(windows)]
    {
        use windows::Win32::Foundation::HANDLE;
        use windows::Win32::Security::{GetTokenInformation, TokenElevation, TOKEN_ELEVATION, TOKEN_QUERY};
        use windows::Win32::System::Threading::{GetCurrentProcess, OpenProcessToken};

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
                )
                .is_ok()
                {
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
