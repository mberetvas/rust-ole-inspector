//! Filter logic for COM objects.
//! 
//! This module contains all filtering and matching logic used during registry scanning.
//! It supports multiple filter types: interactive, description-based, CLSID-based, and app-based.

/// Determines if a COM object should be included based on all active filters
pub fn should_include_object(
    prog_id: &Option<String>,
    description: &Option<String>,
    clsid: &str,
    interactive_filter: &Option<String>,
    filter_description: &Option<String>,
    filter_clsid: &Option<String>,
    filter_app: &Option<Vec<String>>,
) -> bool {
    // Check interactive filter (searches ProgID, description, and CLSID)
    if let Some(ref filter) = interactive_filter {
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
            return false;
        }
    }

    // Check description filter
    if let Some(ref desc_filter) = filter_description {
        let desc_filter_lower = desc_filter.to_lowercase();
        let desc_matches = description
            .as_ref()
            .map(|d| d.to_lowercase().contains(&desc_filter_lower))
            .unwrap_or(false);

        if !desc_matches {
            return false;
        }
    }

    // Check CLSID filter
    if let Some(ref clsid_filter) = filter_clsid {
        let clsid_filter_lower = clsid_filter.to_lowercase();
        if !clsid.to_lowercase().contains(&clsid_filter_lower) {
            return false;
        }
    }

    // Check app filter (comma-separated keywords)
    if let Some(ref app_filters) = filter_app {
        let matches = app_filters.iter().any(|app| {
            let app_lower = app.to_lowercase();
            prog_id
                .as_ref()
                .map(|p| p.to_lowercase().contains(&app_lower))
                .unwrap_or(false)
                || description
                    .as_ref()
                    .map(|d| d.to_lowercase().contains(&app_lower))
                    .unwrap_or(false)
                || clsid.to_lowercase().contains(&app_lower)
        });

        if !matches {
            return false;
        }
    }

    true
}
