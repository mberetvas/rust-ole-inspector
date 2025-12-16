use clap::Parser;

/// A Rust CLI for Windows that discovers COM objects and checks their programmatic usability
#[derive(Parser, Debug)]
#[command(author, version, about, long_about = None)]
pub struct Args {
    /// Show detailed information for each COM object
    #[arg(short, long)]
    pub verbose: bool,

    /// Scan 32-bit registry view
    #[arg(long)]
    pub scan_32bit: bool,

    /// Scan 64-bit registry view (default on 64-bit systems)
    #[arg(long)]
    pub scan_64bit: bool,

    /// Limit the number of results (0 = no limit)
    #[arg(short, long, default_value = "0")]
    pub limit: usize,

    /// Filter by ProgID, description, or CLSID substring (case-insensitive)
    #[arg(short, long)]
    pub filter: Option<String>,

    /// Filter by description substring only (case-insensitive)
    #[arg(long)]
    pub filter_description: Option<String>,

    /// Filter by CLSID substring only (case-insensitive)
    #[arg(long)]
    pub filter_clsid: Option<String>,

    /// Filter by application keywords (comma-separated, case-insensitive)
    #[arg(long, value_delimiter = ',')]
    pub filter_app: Option<Vec<String>>,
}

/// Represents a COM object found in the Windows registry
#[derive(Debug, Clone)]
pub struct ComObject {
    pub clsid: String,
    pub prog_id: Option<String>,
    pub description: Option<String>,
}
