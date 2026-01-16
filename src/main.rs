//! Font Enumerator - A Windows desktop application for exploring system fonts
//!
//! This application demonstrates three different Windows APIs for font enumeration:
//! 1. GDI (Graphics Device Interface) - Legacy API, available on all Windows versions
//! 2. DirectWrite - Modern API with better Unicode support and font metrics
//! 3. FontSet API - Windows 10+ API with access to variable font axes and file paths
//!
//! ## Architecture Overview
//!
//! The application follows a typical Win32 GUI structure:
//! - Single main window with child controls (buttons, listview, preview panel)
//! - Thread-local application state (AppState) to avoid global mutable statics
//! - Message-driven event handling through the window procedure (wnd_proc)
//!
//! ## Code Organization
//!
//! 1. Imports & Constants (lines ~1-65)
//! 2. Data Structures - FontInfo, AppState, EnumMode (lines ~27-65)
//! 3. Entry Point - main() (lines ~67-131)
//! 4. Window Procedure - wnd_proc() handles all window messages (lines ~133-232)
//! 5. UI Creation & Layout - create_controls(), resize_controls() (lines ~234-401)
//! 6. Font Enumeration - GDI, DirectWrite, FontSet implementations (lines ~403-680)
//! 7. String Helpers - DirectWrite string extraction utilities (lines ~682-733)
//! 8. Filtering & Display - apply_filter(), populate_list_view(), etc. (lines ~735-end)

#![windows_subsystem = "windows"]

use std::cell::RefCell;
use std::ffi::c_void;
use windows::{
    core::*,
    Win32::{
        Foundation::*,
        Graphics::DirectWrite::*,
        Graphics::Gdi::*,
        System::LibraryLoader::GetModuleHandleW,
        UI::Controls::*,
        UI::WindowsAndMessaging::*,
    },
};

// ============================================================================
// CONSTANTS - Control IDs for child windows
// ============================================================================
// These IDs are used to identify controls in WM_COMMAND and WM_NOTIFY messages

const IDC_LISTVIEW: u16 = 1001;        // Main font list
const IDC_GDI_BUTTON: u16 = 1002;      // "GDI" enumeration button
const IDC_DWRITE_BUTTON: u16 = 1003;   // "DirectWrite" enumeration button
const IDC_FONTSET_BUTTON: u16 = 1004;  // "FontSet API" enumeration button
const IDC_PREVIEW_STATIC: u16 = 1005;  // Font preview panel
const IDC_STATUS_LABEL: u16 = 1006;    // Status text showing font count
const IDC_SEARCH_EDIT: u16 = 1007;     // Filter text input
const IDC_SEARCH_LABEL: u16 = 1008;    // "Filter:" label

// ============================================================================
// DATA STRUCTURES
// ============================================================================

/// Represents information about a single font face
///
/// Different enumeration APIs provide different levels of detail:
/// - GDI: family_name, style_name, weight, italic, fixed_pitch
/// - DirectWrite: Same as GDI plus better Unicode handling
/// - FontSet: All above plus file_path, variable_axes, is_variable
#[derive(Clone, Default)]
struct FontInfo {
    family_name: String,    // e.g., "Arial", "Segoe UI"
    style_name: String,     // e.g., "Regular", "Bold Italic"
    file_path: String,      // Full path to font file (FontSet API only)
    variable_axes: String,  // Variable font axes, e.g., "wght 100-900" (FontSet API only)
    weight: i32,            // Font weight: 400=Normal, 700=Bold, etc.
    italic: bool,           // Whether this is an italic/oblique style
    fixed_pitch: bool,      // True for monospace fonts
    is_variable: bool,      // True if font has variable axes
}

/// Application state stored in thread-local storage
///
/// Win32 callbacks (like wnd_proc) can't easily access Rust structs,
/// so we use thread_local! with RefCell to provide interior mutability.
#[derive(Default)]
struct AppState {
    // Window handles
    hwnd: HWND,                 // Main window
    h_instance: HINSTANCE,      // Application instance
    list_view: HWND,            // ListView control
    status_label: HWND,         // Status text control
    search_edit: HWND,          // Filter input control
    preview_static: HWND,       // Preview panel control

    // Font data
    fonts: Vec<FontInfo>,           // All enumerated fonts
    filtered_indices: Vec<usize>,   // Indices of fonts matching filter
    filter_text: String,            // Current filter string
    current_mode: EnumMode,         // Which API was used for enumeration
    selected_font: String,          // Currently selected font family
}

/// Enumeration mode - tracks which API was used to enumerate fonts
#[derive(Default, Clone, Copy, PartialEq)]
enum EnumMode {
    #[default]
    None,        // No enumeration performed yet
    Gdi,         // EnumFontFamiliesEx (legacy)
    DirectWrite, // IDWriteFontCollection (modern)
    FontSet,     // IDWriteFontSet (Windows 10+)
}

// Thread-local storage for application state
// This pattern avoids unsafe global mutable statics while allowing
// the window procedure callback to access application data
thread_local! {
    static APP_STATE: RefCell<AppState> = RefCell::new(AppState::default());
}

// ============================================================================
// ENTRY POINT
// ============================================================================

fn main() -> Result<()> {
    unsafe {
        let instance: HINSTANCE = GetModuleHandleW(None)?.into();

        // Initialize common controls (required for ListView)
        let icex = INITCOMMONCONTROLSEX {
            dwSize: std::mem::size_of::<INITCOMMONCONTROLSEX>() as u32,
            dwICC: ICC_LISTVIEW_CLASSES,
        };
        let _ = InitCommonControlsEx(&icex);

        // Register the main window class
        let class_name = w!("FontEnumRustWindowClass");
        let wc = WNDCLASSEXW {
            cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
            style: CS_HREDRAW | CS_VREDRAW,  // Redraw on size change
            lpfnWndProc: Some(wnd_proc),     // Message handler
            hInstance: instance,
            hCursor: LoadCursorW(None, IDC_ARROW)?,
            hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as *mut c_void),
            lpszClassName: class_name,
            hIcon: LoadIconW(None, IDI_APPLICATION)?,
            hIconSm: LoadIconW(None, IDI_APPLICATION)?,
            ..Default::default()
        };

        if RegisterClassExW(&wc) == 0 {
            return Err(Error::from_win32());
        }

        // Create the main window
        let hwnd = CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            class_name,
            w!("Font Enumerator (Rust) - GDI, DirectWrite & FontSet API"),
            WS_OVERLAPPEDWINDOW,
            CW_USEDEFAULT, CW_USEDEFAULT,  // Default position
            1100, 650,                      // Initial size
            HWND::default(),
            HMENU::default(),
            instance,
            None,
        )?;

        // Store handles in app state for later use
        APP_STATE.with(|state| {
            let mut s = state.borrow_mut();
            s.hwnd = hwnd;
            s.h_instance = instance;
        });

        let _ = ShowWindow(hwnd, SW_SHOW);
        let _ = UpdateWindow(hwnd);

        // Standard Win32 message loop
        let mut msg = MSG::default();
        while GetMessageW(&mut msg, None, 0, 0).into() {
            let _ = TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }

        Ok(())
    }
}

// ============================================================================
// WINDOW PROCEDURE - Main message handler
// ============================================================================

/// Handles all window messages for the main window
///
/// Key messages handled:
/// - WM_CREATE: Initialize child controls
/// - WM_SIZE: Resize controls to fit window
/// - WM_COMMAND: Button clicks and edit control changes
/// - WM_NOTIFY: ListView selection changes
/// - WM_DESTROY: Clean up and exit
unsafe extern "system" fn wnd_proc(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    match msg {
        WM_CREATE => {
            create_controls(hwnd);
            LRESULT(0)
        }

        WM_SIZE => {
            resize_controls(hwnd);
            LRESULT(0)
        }

        // Handle button clicks and edit control notifications
        WM_COMMAND => {
            let control_id = (wparam.0 & 0xFFFF) as u16;
            let notification = ((wparam.0 >> 16) & 0xFFFF) as u16;

            match control_id {
                IDC_GDI_BUTTON => enumerate_gdi_fonts(),
                IDC_DWRITE_BUTTON => enumerate_directwrite_fonts(),
                IDC_FONTSET_BUTTON => enumerate_fontset_fonts(),

                // Filter text changed - reapply filter
                IDC_SEARCH_EDIT if notification == EN_CHANGE as u16 => {
                    let mut buffer = [0u16; 256];
                    APP_STATE.with(|state| {
                        let state = state.borrow();
                        let _ = GetWindowTextW(state.search_edit, &mut buffer);
                    });
                    let filter = String::from_utf16_lossy(&buffer)
                        .trim_end_matches('\0')
                        .to_string();
                    APP_STATE.with(|state| {
                        state.borrow_mut().filter_text = filter;
                    });
                    apply_filter();
                }
                _ => {}
            }
            LRESULT(0)
        }

        // Handle ListView notifications (selection changes)
        WM_NOTIFY => {
            let nmhdr = &*(lparam.0 as *const NMHDR);

            // Check if notification is from our ListView
            if nmhdr.idFrom == IDC_LISTVIEW as usize && nmhdr.code == LVN_ITEMCHANGED {
                let nmlv = &*(lparam.0 as *const NMLISTVIEW);

                // Only respond to selection (not deselection)
                if (nmlv.uNewState & LVIS_SELECTED.0) != 0 {
                    // Extract font info from app state
                    let (preview_hwnd, font_name, font_weight, font_italic, style_name) = APP_STATE.with(|state| {
                        let mut state = state.borrow_mut();
                        if let Some(&idx) = state.filtered_indices.get(nmlv.iItem as usize) {
                            if idx < state.fonts.len() {
                                let font = &state.fonts[idx];
                                let family_name = font.family_name.clone();
                                let style_name = font.style_name.clone();
                                let weight = font.weight;
                                let italic = font.italic;
                                state.selected_font = family_name.clone();
                                return (state.preview_static, family_name, weight, italic, style_name);
                            }
                        }
                        (HWND::default(), String::new(), 400, false, String::new())
                    });

                    // Update the preview panel with selected font
                    if preview_hwnd != HWND::default() && !font_name.is_empty() {
                        // Create a font handle with the selected family, weight, and italic
                        let font_name_wide: Vec<u16> = font_name.encode_utf16().chain(std::iter::once(0)).collect();
                        let hfont = CreateFontW(
                            32,  // Height in logical units (pixels at 96 DPI)
                            0, 0, 0,
                            font_weight,                              // Use actual weight (400, 700, etc.)
                            if font_italic { 1 } else { 0 },          // Use actual italic flag
                            0, 0,                                     // No underline/strikeout
                            DEFAULT_CHARSET.0 as u32,
                            OUT_DEFAULT_PRECIS.0 as u32,
                            CLIP_DEFAULT_PRECIS.0 as u32,
                            CLEARTYPE_QUALITY.0 as u32,
                            (DEFAULT_PITCH.0 | FF_DONTCARE.0) as u32,
                            PCWSTR(font_name_wide.as_ptr()),
                        );

                        // Apply the font to the preview control
                        let _ = SendMessageW(preview_hwnd, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));

                        // Set preview text showing font name and sample characters
                        let preview_text = format!(
                            "{} {}\r\n\r\nAaBbCcDdEeFfGgHhIiJjKk\r\n\r\n0123456789 !@#$%",
                            font_name, style_name
                        );
                        let preview_wide: Vec<u16> = preview_text.encode_utf16().chain(std::iter::once(0)).collect();
                        let _ = SetWindowTextW(preview_hwnd, PCWSTR(preview_wide.as_ptr()));
                    }
                }
            }
            LRESULT(0)
        }

        // Set minimum window size
        WM_GETMINMAXINFO => {
            let mmi = &mut *(lparam.0 as *mut MINMAXINFO);
            mmi.ptMinTrackSize.x = 800;
            mmi.ptMinTrackSize.y = 500;
            LRESULT(0)
        }

        WM_DESTROY => {
            PostQuitMessage(0);
            LRESULT(0)
        }

        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

// ============================================================================
// UI CREATION & LAYOUT
// ============================================================================

/// Creates all child controls for the main window
///
/// Layout:
/// ```text
/// +------------------------------------------------------------------+
/// | [GDI] [DirectWrite] [FontSet API]  Filter: [____]  Status text   |
/// +--------------------------------+--------------------------------+
/// |                                |                                 |
/// |         ListView               |        Preview Panel            |
/// |     (font list table)          |    (sample text in font)        |
/// |                                |                                 |
/// +--------------------------------+---------------------------------+
/// ```
unsafe fn create_controls(hwnd: HWND) {
    let instance = APP_STATE.with(|state| state.borrow().h_instance);

    // --- Toolbar buttons ---
    let _ = CreateWindowExW(
        WINDOW_EX_STYLE::default(),
        w!("BUTTON"),
        w!("GDI"),
        WS_CHILD | WS_VISIBLE | WINDOW_STYLE(BS_PUSHBUTTON as u32),
        10, 10, 80, 30,
        hwnd,
        HMENU(IDC_GDI_BUTTON as *mut c_void),
        instance,
        None,
    );

    let _ = CreateWindowExW(
        WINDOW_EX_STYLE::default(),
        w!("BUTTON"),
        w!("DirectWrite"),
        WS_CHILD | WS_VISIBLE | WINDOW_STYLE(BS_PUSHBUTTON as u32),
        100, 10, 100, 30,
        hwnd,
        HMENU(IDC_DWRITE_BUTTON as *mut c_void),
        instance,
        None,
    );

    let _ = CreateWindowExW(
        WINDOW_EX_STYLE::default(),
        w!("BUTTON"),
        w!("FontSet API"),
        WS_CHILD | WS_VISIBLE | WINDOW_STYLE(BS_PUSHBUTTON as u32),
        210, 10, 100, 30,
        hwnd,
        HMENU(IDC_FONTSET_BUTTON as *mut c_void),
        instance,
        None,
    );

    // --- Filter controls ---
    let _ = CreateWindowExW(
        WINDOW_EX_STYLE::default(),
        w!("STATIC"),
        w!("Filter:"),
        WS_CHILD | WS_VISIBLE,
        330, 17, 40, 20,
        hwnd,
        HMENU(IDC_SEARCH_LABEL as *mut c_void),
        instance,
        None,
    );

    let search_edit = CreateWindowExW(
        WS_EX_CLIENTEDGE,  // Sunken edge style
        w!("EDIT"),
        w!(""),
        WS_CHILD | WS_VISIBLE | WINDOW_STYLE(ES_AUTOHSCROLL as u32),
        375, 12, 180, 24,
        hwnd,
        HMENU(IDC_SEARCH_EDIT as *mut c_void),
        instance,
        None,
    ).unwrap_or_default();

    // --- Status label ---
    let status_label = CreateWindowExW(
        WINDOW_EX_STYLE::default(),
        w!("STATIC"),
        w!("Click a button to enumerate fonts"),
        WS_CHILD | WS_VISIBLE,
        570, 17, 350, 20,
        hwnd,
        HMENU(IDC_STATUS_LABEL as *mut c_void),
        instance,
        None,
    ).unwrap_or_default();

    // --- ListView (font list) ---
    let list_view = CreateWindowExW(
        WS_EX_CLIENTEDGE,
        w!("SysListView32"),
        w!(""),
        WS_CHILD | WS_VISIBLE | WINDOW_STYLE((LVS_REPORT | LVS_SINGLESEL | LVS_SHOWSELALWAYS) as u32),
        10, 50, 600, 500,
        hwnd,
        HMENU(IDC_LISTVIEW as *mut c_void),
        instance,
        None,
    ).unwrap_or_default();

    // Enable modern ListView features
    let _ = SendMessageW(
        list_view,
        LVM_SETEXTENDEDLISTVIEWSTYLE,
        WPARAM(0),
        LPARAM((LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_DOUBLEBUFFER) as isize),
    );

    // Add columns to ListView
    add_column(list_view, 0, "Font Family", 180);
    add_column(list_view, 1, "Style", 100);
    add_column(list_view, 2, "Weight", 60);
    add_column(list_view, 3, "Italic", 50);
    add_column(list_view, 4, "Fixed", 50);
    add_column(list_view, 5, "File Path", 180);
    add_column(list_view, 6, "Variable Axes", 180);

    // --- Preview panel ---
    // Using multiline EDIT control (read-only) for easy font display
    // ES_MULTILINE = 0x0004, ES_READONLY = 0x0800
    let preview_static = CreateWindowExW(
        WS_EX_CLIENTEDGE,
        w!("EDIT"),
        w!("Select a font to preview"),
        WS_CHILD | WS_VISIBLE | WINDOW_STYLE(0x0004 | 0x0800),
        620, 50, 350, 400,
        hwnd,
        HMENU(IDC_PREVIEW_STATIC as *mut c_void),
        instance,
        None,
    ).unwrap_or_default();

    // Store control handles in app state
    APP_STATE.with(|state| {
        let mut state = state.borrow_mut();
        state.list_view = list_view;
        state.status_label = status_label;
        state.search_edit = search_edit;
        state.preview_static = preview_static;
    });
}

/// Helper function to add a column to the ListView
unsafe fn add_column(list_view: HWND, index: i32, text: &str, width: i32) {
    let text_wide: Vec<u16> = text.encode_utf16().chain(std::iter::once(0)).collect();
    let col = LVCOLUMNW {
        mask: LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM,
        cx: width,
        pszText: PWSTR(text_wide.as_ptr() as *mut u16),
        iSubItem: index,
        ..Default::default()
    };
    let _ = SendMessageW(
        list_view,
        LVM_INSERTCOLUMNW,
        WPARAM(index as usize),
        LPARAM(&col as *const _ as isize),
    );
}

/// Resizes child controls when the window size changes
///
/// The layout splits the content area 60/40 between the list and preview panel.
unsafe fn resize_controls(hwnd: HWND) {
    let mut rect = RECT::default();
    let _ = GetClientRect(hwnd, &mut rect);

    let width = rect.right - rect.left;
    let height = rect.bottom - rect.top;
    let list_height = height - 70;  // Leave space for toolbar

    // Calculate widths - 60% for list, 40% for preview
    let list_w = ((width - 40) * 60) / 100;
    let preview_x = list_w + 20;
    let preview_w = width - preview_x - 10;

    APP_STATE.with(|state| {
        let state = state.borrow();
        let _ = MoveWindow(state.list_view, 10, 50, list_w, list_height, true);
        let _ = MoveWindow(state.preview_static, preview_x, 50, preview_w, list_height, true);
    });
}

// ============================================================================
// FONT ENUMERATION - GDI API
// ============================================================================

/// Callback function for GDI font enumeration
///
/// Called once for each font face found by EnumFontFamiliesExW.
/// Extracts font information and adds unique fonts to the collection.
unsafe extern "system" fn enum_font_proc(
    lpelfe: *const LOGFONTW,
    _lpntme: *const TEXTMETRICW,
    _font_type: u32,
    lparam: LPARAM,
) -> i32 {
    let fonts = &mut *(lparam.0 as *mut Vec<FontInfo>);
    let lf = &*lpelfe;
    let elfex = &*(lpelfe as *const ENUMLOGFONTEXW);

    // Extract font names from wide strings
    let family_name = String::from_utf16_lossy(&lf.lfFaceName)
        .trim_end_matches('\0')
        .to_string();
    let style_name = String::from_utf16_lossy(&elfex.elfStyle)
        .trim_end_matches('\0')
        .to_string();

    // Skip duplicates (same family + style)
    let exists = fonts.iter().any(|f| f.family_name == family_name && f.style_name == style_name);

    if !exists {
        // Check if font is fixed-pitch (monospace)
        // FIXED_PITCH is value 1 in the low 2 bits of lfPitchAndFamily
        let pitch_and_family: u8 = std::mem::transmute(lf.lfPitchAndFamily);
        let is_fixed = (pitch_and_family & 0x03) == 1;

        fonts.push(FontInfo {
            family_name,
            style_name,
            weight: lf.lfWeight,
            italic: lf.lfItalic != 0,
            fixed_pitch: is_fixed,
            ..Default::default()
        });
    }

    1 // Return 1 to continue enumeration
}

/// Enumerates fonts using the GDI EnumFontFamiliesEx API
///
/// This is the oldest font enumeration API, available on all Windows versions.
/// Limitations:
/// - No access to font file paths
/// - No variable font axis information
/// - Limited style name accuracy for some fonts
fn enumerate_gdi_fonts() {
    unsafe {
        let mut fonts: Vec<FontInfo> = Vec::new();

        APP_STATE.with(|state| {
            let state = state.borrow();
            let hdc = GetDC(state.hwnd);

            // Set up LOGFONT to enumerate all fonts (DEFAULT_CHARSET = 1)
            let mut lf = LOGFONTW {
                lfCharSet: FONT_CHARSET(1),
                ..Default::default()
            };

            // Enumerate all font families
            let _ = EnumFontFamiliesExW(
                hdc,
                &mut lf,
                Some(enum_font_proc),
                LPARAM(&mut fonts as *mut _ as isize),
                0,
            );

            let _ = ReleaseDC(state.hwnd, hdc);
        });

        // Sort alphabetically by family name
        fonts.sort_by(|a, b| a.family_name.cmp(&b.family_name));

        // Update app state with enumerated fonts
        APP_STATE.with(|state| {
            let mut state = state.borrow_mut();
            state.fonts = fonts;
            state.current_mode = EnumMode::Gdi;
            state.selected_font.clear();
        });

        apply_filter();
    }
}

// ============================================================================
// FONT ENUMERATION - DirectWrite API
// ============================================================================

/// Enumerates fonts using the DirectWrite IDWriteFontCollection API
///
/// DirectWrite provides better support for:
/// - OpenType features
/// - Complex script shaping
/// - Font fallback
/// - Accurate style names
///
/// Available on Windows Vista and later.
fn enumerate_directwrite_fonts() {
    unsafe {
        let mut fonts: Vec<FontInfo> = Vec::new();

        // Create DirectWrite factory
        let factory: IDWriteFactory = match DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED) {
            Ok(f) => f,
            Err(_) => return,
        };

        // Get the system font collection
        let mut collection: Option<IDWriteFontCollection> = None;
        if factory.GetSystemFontCollection(&mut collection, false).is_err() {
            return;
        }
        let collection = match collection {
            Some(c) => c,
            None => return,
        };

        let family_count = collection.GetFontFamilyCount();

        // Iterate through each font family
        for i in 0..family_count {
            if let Ok(family) = collection.GetFontFamily(i) {
                let family_name = get_family_names(&family);

                // Each family can contain multiple fonts (Regular, Bold, Italic, etc.)
                let font_count = family.GetFontCount();
                for j in 0..font_count {
                    if let Ok(font) = family.GetFont(j) {
                        let style_name = get_face_names(&font);

                        // Check if font is monospaced (requires IDWriteFont1)
                        let is_mono = font
                            .cast::<IDWriteFont1>()
                            .map(|f1| f1.IsMonospacedFont().as_bool())
                            .unwrap_or(false);

                        fonts.push(FontInfo {
                            family_name: family_name.clone(),
                            style_name,
                            weight: font.GetWeight().0 as i32,
                            italic: font.GetStyle() != DWRITE_FONT_STYLE_NORMAL,
                            fixed_pitch: is_mono,
                            ..Default::default()
                        });
                    }
                }
            }
        }

        // Sort by family name, then by style name
        fonts.sort_by(|a, b| {
            a.family_name
                .cmp(&b.family_name)
                .then(a.style_name.cmp(&b.style_name))
        });

        APP_STATE.with(|state| {
            let mut state = state.borrow_mut();
            state.fonts = fonts;
            state.current_mode = EnumMode::DirectWrite;
            state.selected_font.clear();
        });

        apply_filter();
    }
}

// ============================================================================
// FONT ENUMERATION - FontSet API (Windows 10+)
// ============================================================================

/// Enumerates fonts using the DirectWrite IDWriteFontSet API
///
/// The FontSet API (Windows 10+) provides access to:
/// - Font file paths
/// - Variable font axis information (weight ranges, width ranges, etc.)
/// - More detailed font properties
///
/// This is the most comprehensive font enumeration API available.
fn enumerate_fontset_fonts() {
    unsafe {
        let mut fonts: Vec<FontInfo> = Vec::new();

        // Create DirectWrite factory (version 3 required for FontSet API)
        let factory: IDWriteFactory3 = match DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED) {
            Ok(f) => f,
            Err(_) => {
                APP_STATE.with(|state| {
                    let state = state.borrow();
                    let _ = MessageBoxW(
                        state.hwnd,
                        w!("Failed to create DirectWrite factory 3.\nThis feature requires Windows 10 or later."),
                        w!("Error"),
                        MB_OK | MB_ICONERROR,
                    );
                });
                return;
            }
        };

        // Get the system font set
        let font_set: IDWriteFontSet = match factory.GetSystemFontSet() {
            Ok(fs) => fs,
            Err(_) => return,
        };

        let font_count = font_set.GetFontCount();

        // Iterate through each font in the set
        for i in 0..font_count {
            let mut info = FontInfo::default();

            // Get font face reference for accessing file info
            if let Ok(font_ref) = font_set.GetFontFaceReference(i) {
                // --- Extract font file path ---
                if let Ok(font_file) = font_ref.GetFontFile() {
                    if let Ok(loader) = font_file.GetLoader() {
                        // Only local fonts have file paths
                        if let Ok(local_loader) = loader.cast::<IDWriteLocalFontFileLoader>() {
                            let mut ref_key: *mut c_void = std::ptr::null_mut();
                            let mut ref_key_size: u32 = 0;
                            if font_file.GetReferenceKey(&mut ref_key, &mut ref_key_size).is_ok() {
                                if let Ok(path_len) = local_loader.GetFilePathLengthFromKey(ref_key, ref_key_size) {
                                    let mut path_buf = vec![0u16; (path_len + 1) as usize];
                                    if local_loader.GetFilePathFromKey(ref_key, ref_key_size, &mut path_buf).is_ok() {
                                        info.file_path = String::from_utf16_lossy(&path_buf)
                                            .trim_end_matches('\0')
                                            .to_string();
                                    }
                                }
                            }
                        }
                    }
                }

                // --- Extract variable font axis information ---
                if let Ok(font_face3) = font_ref.CreateFontFace() {
                    if let Ok(font_face5) = font_face3.cast::<IDWriteFontFace5>() {
                        if let Ok(font_resource) = font_face5.GetFontResource() {
                            let axis_count = font_resource.GetFontAxisCount();
                            if axis_count > 0 {
                                let mut axis_ranges = vec![DWRITE_FONT_AXIS_RANGE::default(); axis_count as usize];
                                if font_resource.GetFontAxisRanges(&mut axis_ranges).is_ok() {
                                    for range in &axis_ranges {
                                        // Variable axis has different min/max values
                                        if range.minValue != range.maxValue {
                                            info.is_variable = true;
                                            if !info.variable_axes.is_empty() {
                                                info.variable_axes.push_str(", ");
                                            }
                                            // Convert 4-byte axis tag to string (e.g., "wght", "wdth")
                                            let tag = range.axisTag.0;
                                            let tag_str = format!(
                                                "{}{}{}{}",
                                                (tag & 0xFF) as u8 as char,
                                                ((tag >> 8) & 0xFF) as u8 as char,
                                                ((tag >> 16) & 0xFF) as u8 as char,
                                                ((tag >> 24) & 0xFF) as u8 as char
                                            );
                                            info.variable_axes.push_str(&format!(
                                                "{} {}-{}",
                                                tag_str,
                                                range.minValue as i32,
                                                range.maxValue as i32
                                            ));
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // --- Extract font properties from the font set ---
            if let Ok(prop) = font_set.GetPropertyValues(DWRITE_FONT_PROPERTY_ID_FAMILY_NAME) {
                if i < prop.GetCount() {
                    info.family_name = get_string_from_string_list(&prop, i);
                }
            }

            if let Ok(prop) = font_set.GetPropertyValues(DWRITE_FONT_PROPERTY_ID_FACE_NAME) {
                if i < prop.GetCount() {
                    info.style_name = get_string_from_string_list(&prop, i);
                }
            }

            if let Ok(prop) = font_set.GetPropertyValues(DWRITE_FONT_PROPERTY_ID_WEIGHT) {
                if i < prop.GetCount() {
                    let s = get_string_from_string_list(&prop, i);
                    info.weight = s.parse().unwrap_or(400);
                }
            }

            if let Ok(prop) = font_set.GetPropertyValues(DWRITE_FONT_PROPERTY_ID_STYLE) {
                if i < prop.GetCount() {
                    let s = get_string_from_string_list(&prop, i);
                    let style: i32 = s.parse().unwrap_or(0);
                    info.italic = style != 0;  // 0 = normal, 1 = italic, 2 = oblique
                }
            }

            if !info.family_name.is_empty() {
                fonts.push(info);
            }
        }

        fonts.sort_by(|a, b| {
            a.family_name
                .cmp(&b.family_name)
                .then(a.style_name.cmp(&b.style_name))
        });

        APP_STATE.with(|state| {
            let mut state = state.borrow_mut();
            state.fonts = fonts;
            state.current_mode = EnumMode::FontSet;
            state.selected_font.clear();
        });

        apply_filter();
    }
}

// ============================================================================
// DIRECTWRITE STRING HELPERS
// ============================================================================

/// Extracts the family name from a DirectWrite font family
fn get_family_names(family: &IDWriteFontFamily) -> String {
    unsafe {
        if let Ok(names) = family.GetFamilyNames() {
            return get_string_from_localized(&names);
        }
        String::new()
    }
}

/// Extracts the face/style name from a DirectWrite font
fn get_face_names(font: &IDWriteFont) -> String {
    unsafe {
        if let Ok(names) = font.GetFaceNames() {
            return get_string_from_localized(&names);
        }
        String::new()
    }
}

/// Extracts a string from IDWriteLocalizedStrings, preferring English
fn get_string_from_localized(strings: &IDWriteLocalizedStrings) -> String {
    unsafe {
        let mut index: u32 = 0;
        let mut exists = BOOL::default();

        // Try to find English (US) version first
        let _ = strings.FindLocaleName(w!("en-us"), &mut index, &mut exists);
        if !exists.as_bool() {
            index = 0;  // Fall back to first available
        }

        if let Ok(length) = strings.GetStringLength(index) {
            let mut buffer = vec![0u16; (length + 1) as usize];
            if strings.GetString(index, &mut buffer).is_ok() {
                return String::from_utf16_lossy(&buffer)
                    .trim_end_matches('\0')
                    .to_string();
            }
        }
        String::new()
    }
}

/// Extracts a string from IDWriteStringList by index
fn get_string_from_string_list(strings: &IDWriteStringList, index: u32) -> String {
    unsafe {
        if let Ok(length) = strings.GetStringLength(index) {
            let mut buffer = vec![0u16; (length + 1) as usize];
            if strings.GetString(index, &mut buffer).is_ok() {
                return String::from_utf16_lossy(&buffer)
                    .trim_end_matches('\0')
                    .to_string();
            }
        }
        String::new()
    }
}

// ============================================================================
// FILTERING & DISPLAY
// ============================================================================

/// Applies the current filter text to the font list
///
/// Creates a list of indices into the fonts vector for fonts that match
/// the filter (case-insensitive search in family name or style name).
fn apply_filter() {
    // Collect data needed for filtering (avoid holding borrow during iteration)
    let (fonts_data, filter_lower): (Vec<(String, String)>, String) = APP_STATE.with(|state| {
        let state = state.borrow();
        let fonts_data: Vec<(String, String)> = state.fonts.iter()
            .map(|f| (f.family_name.clone(), f.style_name.clone()))
            .collect();
        (fonts_data, state.filter_text.to_lowercase())
    });

    // Filter fonts by checking if family or style contains the filter text
    let indices: Vec<usize> = fonts_data.iter().enumerate()
        .filter(|(_, (family, style))| {
            filter_lower.is_empty()
                || family.to_lowercase().contains(&filter_lower)
                || style.to_lowercase().contains(&filter_lower)
        })
        .map(|(i, _)| i)
        .collect();

    APP_STATE.with(|state| {
        state.borrow_mut().filtered_indices = indices;
    });

    populate_list_view();
    update_status_text();

    // Invalidate preview to clear selection
    unsafe {
        APP_STATE.with(|state| {
            let state = state.borrow();
            let _ = InvalidateRect(state.preview_static, None, true);
        });
    }
}

/// Populates the ListView with filtered font data
fn populate_list_view() {
    unsafe {
        APP_STATE.with(|state| {
            let state = state.borrow();

            // Clear existing items
            let _ = SendMessageW(state.list_view, LVM_DELETEALLITEMS, WPARAM(0), LPARAM(0));

            // Add each filtered font to the list
            for (i, &font_idx) in state.filtered_indices.iter().enumerate() {
                let font = &state.fonts[font_idx];

                // Insert main item (family name)
                let family_wide: Vec<u16> = font.family_name.encode_utf16().chain(std::iter::once(0)).collect();
                let item = LVITEMW {
                    mask: LVIF_TEXT | LVIF_PARAM,
                    iItem: i as i32,
                    iSubItem: 0,
                    pszText: PWSTR(family_wide.as_ptr() as *mut u16),
                    lParam: LPARAM(font_idx as isize),  // Store original index for selection handling
                    ..Default::default()
                };
                let _ = SendMessageW(
                    state.list_view,
                    LVM_INSERTITEMW,
                    WPARAM(0),
                    LPARAM(&item as *const _ as isize),
                );

                // Set subitem columns
                set_list_item_text(state.list_view, i as i32, 1, &font.style_name);
                set_list_item_text(state.list_view, i as i32, 2, &font.weight.to_string());
                set_list_item_text(state.list_view, i as i32, 3, if font.italic { "Yes" } else { "No" });
                set_list_item_text(state.list_view, i as i32, 4, if font.fixed_pitch { "Yes" } else { "No" });
                set_list_item_text(state.list_view, i as i32, 5, &font.file_path);

                let var_str = if font.is_variable {
                    format!("Yes: {}", font.variable_axes)
                } else {
                    String::new()
                };
                set_list_item_text(state.list_view, i as i32, 6, &var_str);
            }
        });
    }
}

/// Helper to set text for a ListView subitem
unsafe fn set_list_item_text(list_view: HWND, item: i32, subitem: i32, text: &str) {
    let text_wide: Vec<u16> = text.encode_utf16().chain(std::iter::once(0)).collect();
    let lvi = LVITEMW {
        iSubItem: subitem,
        pszText: PWSTR(text_wide.as_ptr() as *mut u16),
        ..Default::default()
    };
    let _ = SendMessageW(
        list_view,
        LVM_SETITEMTEXTW,
        WPARAM(item as usize),
        LPARAM(&lvi as *const _ as isize),
    );
}

/// Updates the status label with current font count
fn update_status_text() {
    unsafe {
        APP_STATE.with(|state| {
            let state = state.borrow();

            let mode_str = match state.current_mode {
                EnumMode::Gdi => "GDI",
                EnumMode::DirectWrite => "DirectWrite",
                EnumMode::FontSet => "FontSet",
                EnumMode::None => "No",
            };

            let status = if state.filter_text.is_empty() {
                format!("{} Enumeration: Found {} fonts", mode_str, state.fonts.len())
            } else {
                format!(
                    "{} Enumeration: Showing {} of {} fonts",
                    mode_str,
                    state.filtered_indices.len(),
                    state.fonts.len()
                )
            };

            let status_wide: Vec<u16> = status.encode_utf16().chain(std::iter::once(0)).collect();
            let _ = SetWindowTextW(state.status_label, PCWSTR(status_wide.as_ptr()));
        });
    }
}
