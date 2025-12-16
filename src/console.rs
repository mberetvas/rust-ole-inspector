//! Console utilities for Windows terminal output.
//! 
//! This module handles UTF-8 console setup and ASCII/Unicode header art display.

/// Try to enable UTF-8 output in the Windows console. Returns true on success.
pub fn init_console_utf8() -> bool {
    #[cfg(windows)]
    {
        use windows::Win32::System::Console::{SetConsoleCP, SetConsoleOutputCP};

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
pub fn print_header_art_unicode() {
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
pub fn print_header_art_ascii() {
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
