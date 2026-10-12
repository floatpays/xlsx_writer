use rust_xlsxwriter::{
    Color, DocProperties, ExcelDateTime, Format, FormatAlign, FormatBorder, FormatPattern,
    FormatScript, FormatUnderline, Formula, Image, Note, Url, Workbook, Worksheet, XlsxError,
};
use rustler::{Binary, NifTaggedEnum, ResourceArc};
use std::path::Path;
use std::sync::Mutex;

#[derive(NifTaggedEnum, PartialEq)]
enum CellAlignPos {
    Center,
    Left,
    Right,
}

#[derive(NifTaggedEnum, PartialEq)]
enum CellVAlignPos {
    Top,
    Center,
    Bottom,
    Justify,
    Distributed,
}

#[derive(NifTaggedEnum, PartialEq)]
enum CellPattern {
    Solid,
    None,
    Gray125,
    Gray0625,
}

#[derive(NifTaggedEnum, PartialEq)]
enum UnderlineStyle {
    Single,
    Double,
    SingleAccounting,
    DoubleAccounting,
}

#[derive(NifTaggedEnum, PartialEq)]
enum BorderStyle {
    Thin,
    Medium,
    Thick,
    Dashed,
    Dotted,
    Double,
    Hair,
    MediumDashed,
    DashDot,
    MediumDashDot,
    DashDotDot,
    MediumDashDotDot,
    SlantDashDot,
}

#[derive(NifTaggedEnum, PartialEq)]
enum CellFormat {
    Bold,
    Align(CellAlignPos),
    // Examples of numeric formats: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Format.html#examples-2
    NumFormat(String),
    BgColor(String),
    Pattern(CellPattern),
    FontColor(String),
    Italic,
    Underline(UnderlineStyle),
    Strikethrough,
    FontSize(u16),
    FontName(String),
    Superscript,
    Subscript,
    Border(BorderStyle),
    BorderTop(BorderStyle),
    BorderBottom(BorderStyle),
    BorderLeft(BorderStyle),
    BorderRight(BorderStyle),
    BorderColor(String),
    BorderTopColor(String),
    BorderBottomColor(String),
    BorderLeftColor(String),
    BorderRightColor(String),
    TextWrap,
    Valign(CellVAlignPos),
    Rotation(i16),
    Shrink,
    Indent(u8),
}

#[derive(rustler::NifStruct)]
#[module = "XlsxWriter.NoteOptions"]
struct NoteOptions {
    author: Option<String>,
    visible: Option<bool>,
    width: Option<u32>,
    height: Option<u32>,
}

#[derive(NifTaggedEnum)]
enum CellData<'a> {
    Float(f64),
    String(String),
    StringWithFormat(String, Vec<CellFormat>),
    NumberWithFormat(f64, Vec<CellFormat>),
    ImagePath(String),
    Image(Binary<'a>),
    Date(String),
    DateWithFormat(String, Vec<CellFormat>),
    DateTime(String),
    DateTimeWithFormat(String, Vec<CellFormat>),
    Formula(String),
    FormulaWithFormat(String, Vec<CellFormat>),
    Boolean(bool),
    BooleanWithFormat(bool, Vec<CellFormat>),
    Url(String),
    UrlWithText(String, String),
    UrlWithFormat(String, Vec<CellFormat>),
    UrlWithTextAndFormat(String, String, Vec<CellFormat>),
    Blank(Vec<CellFormat>),
    RichString(Vec<(String, Vec<CellFormat>)>),
    RichStringWithFormat(Vec<(String, Vec<CellFormat>)>, Vec<CellFormat>),
}

#[derive(NifTaggedEnum)]
enum Sheet<'a> {
    Write(u32, u16, CellData<'a>),
    SetColumnWidth(u16, u32),
    SetRowHeight(u32, u16),
    SetColumnRangeWidth(u16, u16, u32),
    SetRowRangeHeight(u32, u32, u16),
    SetFreezePanes(u32, u16),
    SetRowHidden(u32),
    SetColumnHidden(u16),
    SetAutofilter(u32, u16, u32, u16),
    MergeRange(u32, u16, u32, u16, CellData<'a>),
    InsertNote(u32, u16, String, NoteOptions),
    SetTabColor(String),
    SetAutofit,
}

#[derive(rustler::NifStruct)]
#[module = "XlsxWriter.WorkbookProperties"]
struct WorkbookProperties {
    author: Option<String>,
    title: Option<String>,
    subject: Option<String>,
    manager: Option<String>,
    company: Option<String>,
    category: Option<String>,
    keywords: Option<String>,
    comment: Option<String>,
    status: Option<String>,
}

struct StreamWorkbook {
    state: Mutex<Option<StreamState>>,
}

struct StreamState {
    workbook: Workbook,
    current_sheet: Option<usize>,
    last_row: Option<u32>,
}

#[rustler::resource_impl]
impl rustler::Resource for StreamWorkbook {}

#[rustler::nif(schedule = "DirtyIo")]
fn stream_open(
    tempdir: Option<String>,
    properties: Option<WorkbookProperties>,
) -> Result<ResourceArc<StreamWorkbook>, String> {
    let mut workbook = Workbook::new();
    if let Some(dir) = tempdir {
        workbook.set_tempdir(dir).map_err(|e| e.to_string())?;
    }
    if let Some(props) = properties {
        set_workbook_properties(&mut workbook, props);
    }
    Ok(ResourceArc::new(StreamWorkbook {
        state: Mutex::new(Some(StreamState {
            workbook,
            current_sheet: None,
            last_row: None,
        })),
    }))
}

#[rustler::nif(schedule = "DirtyIo")]
fn stream_add_sheet(resource: ResourceArc<StreamWorkbook>, name: String) -> Result<(), String> {
    let mut guard = resource.state.lock().map_err(|e| e.to_string())?;
    let state = guard.as_mut().ok_or("Workbook is closed")?;
    let index = state.workbook.worksheets().len();
    let result = state
        .workbook
        .add_worksheet_with_constant_memory()
        .set_name(name)
        .map(|_| ());
    if let Err(error) = result {
        *guard = None;
        return Err(error.to_string());
    }
    state.current_sheet = Some(index);
    state.last_row = None;
    Ok(())
}

#[rustler::nif(schedule = "DirtyIo")]
fn stream_configure_sheet<'a>(
    resource: ResourceArc<StreamWorkbook>,
    instructions: Vec<Sheet<'a>>,
) -> Result<(), String> {
    let mut guard = resource.state.lock().map_err(|e| e.to_string())?;
    let state = guard.as_mut().ok_or("Workbook is closed")?;
    let index = state
        .current_sheet
        .ok_or("Add a sheet before configuring it")?;
    if state.last_row.is_some() {
        return Err("Configure the sheet before writing rows".to_string());
    }

    for instruction in instructions {
        let worksheet = state
            .workbook
            .worksheet_from_index(index)
            .map_err(|e| e.to_string())?;
        let result = match instruction {
            Sheet::SetColumnWidth(col, width) => worksheet.set_column_width(col, width).map(|_| ()),
            Sheet::SetColumnRangeWidth(first, last, width) => {
                let mut result = Ok(());
                for col in first..=last {
                    if let Err(error) = worksheet.set_column_width(col, width) {
                        result = Err(error);
                        break;
                    }
                }
                result
            }
            Sheet::SetFreezePanes(row, col) => worksheet.set_freeze_panes(row, col).map(|_| ()),
            Sheet::SetColumnHidden(col) => worksheet.set_column_hidden(col).map(|_| ()),
            Sheet::SetAutofilter(first_row, first_col, last_row, last_col) => worksheet
                .autofilter(first_row, first_col, last_row, last_col)
                .map(|_| ()),
            Sheet::SetTabColor(color_hex) => {
                let color = parse_hex_color(&color_hex).ok_or("Invalid tab color")?;
                worksheet.set_tab_color(color);
                Ok(())
            }
            _ => return Err("Unsupported sheet instruction in constant memory mode".to_string()),
        };
        if let Err(error) = result {
            return Err(error.to_string());
        }
    }
    Ok(())
}

#[rustler::nif(schedule = "DirtyIo")]
fn stream_write_cells<'a>(
    resource: ResourceArc<StreamWorkbook>,
    cells: Vec<(u32, u16, CellData<'a>)>,
) -> Result<(), String> {
    let mut guard = resource.state.lock().map_err(|e| e.to_string())?;
    let state = guard.as_mut().ok_or("Workbook is closed")?;
    let index = state
        .current_sheet
        .ok_or("Add a sheet before writing cells")?;
    let mut last_row = state.last_row;

    for (row, col, data) in cells {
        if last_row.is_some_and(|previous| row < previous) {
            *guard = None;
            return Err(format!(
                "Row {row} follows row {} in constant memory mode",
                last_row.unwrap()
            ));
        }
        let result = state
            .workbook
            .worksheet_from_index(index)
            .and_then(|worksheet| write_data(worksheet, row, col, data).map(|_| ()));
        if let Err(error) = result {
            *guard = None;
            return Err(error.to_string());
        }
        last_row = Some(row);
    }
    state.last_row = last_row;
    Ok(())
}

#[rustler::nif(schedule = "DirtyIo")]
fn stream_finish(resource: ResourceArc<StreamWorkbook>, path: String) -> Result<(), String> {
    let mut guard = resource.state.lock().map_err(|e| e.to_string())?;
    let mut state = guard.take().ok_or("Workbook is closed")?;
    if state.current_sheet.is_none() {
        return Err("Add at least one sheet before finishing".to_string());
    }

    let output = Path::new(&path);
    let parent = output
        .parent()
        .filter(|p| !p.as_os_str().is_empty())
        .unwrap_or(Path::new("."));
    let temporary = tempfile::Builder::new()
        .prefix(".xlsx_writer-")
        .tempfile_in(parent)
        .map_err(|e| e.to_string())?
        .into_temp_path();
    state.workbook.save(&temporary).map_err(|e| e.to_string())?;
    temporary.persist(output).map_err(|e| e.to_string())?;
    Ok(())
}

#[rustler::nif]
fn stream_abort(resource: ResourceArc<StreamWorkbook>) -> Result<(), String> {
    let mut guard = resource.state.lock().map_err(|e| e.to_string())?;
    *guard = None;
    Ok(())
}

#[rustler::nif(schedule = "DirtyCpu")]
fn write(sheets: Vec<(String, Vec<Sheet>)>) -> Result<Vec<u8>, String> {
    write_impl(sheets, None)
}

#[rustler::nif(schedule = "DirtyCpu")]
fn write_with_properties(
    sheets: Vec<(String, Vec<Sheet>)>,
    properties: WorkbookProperties,
) -> Result<Vec<u8>, String> {
    write_impl(sheets, Some(properties))
}

fn write_impl(
    sheets: Vec<(String, Vec<Sheet>)>,
    properties: Option<WorkbookProperties>,
) -> Result<Vec<u8>, String> {
    let mut workbook = Workbook::new();

    if let Some(props) = properties {
        set_workbook_properties(&mut workbook, props);
    }

    for (sheet_name, sheet) in sheets {
        let mut worksheet = workbook.add_worksheet();

        if let Err(e) = worksheet.set_name(sheet_name) {
            return Err(e.to_string());
        }

        for instruction in sheet {
            worksheet = match instruction {
                Sheet::SetColumnWidth(col, val) => match worksheet.set_column_width(col, val) {
                    Ok(ws) => ws,
                    Err(e) => return Err(e.to_string()),
                },
                Sheet::SetRowHeight(row, val) => match worksheet.set_row_height(row, val) {
                    Ok(ws) => ws,
                    Err(e) => return Err(e.to_string()),
                },
                Sheet::SetColumnRangeWidth(first_col, last_col, width) => {
                    let mut ws = worksheet;
                    for col in first_col..=last_col {
                        ws = match ws.set_column_width(col, width) {
                            Ok(w) => w,
                            Err(e) => return Err(e.to_string()),
                        };
                    }
                    ws
                }
                Sheet::SetRowRangeHeight(first_row, last_row, height) => {
                    let mut ws = worksheet;
                    for row in first_row..=last_row {
                        ws = match ws.set_row_height(row, height) {
                            Ok(w) => w,
                            Err(e) => return Err(e.to_string()),
                        };
                    }
                    ws
                }
                Sheet::SetFreezePanes(row, col) => match worksheet.set_freeze_panes(row, col) {
                    Ok(ws) => ws,
                    Err(e) => return Err(e.to_string()),
                },
                Sheet::SetRowHidden(row) => match worksheet.set_row_hidden(row) {
                    Ok(ws) => ws,
                    Err(e) => return Err(e.to_string()),
                },
                Sheet::SetColumnHidden(col) => match worksheet.set_column_hidden(col) {
                    Ok(ws) => ws,
                    Err(e) => return Err(e.to_string()),
                },
                Sheet::SetAutofilter(first_row, first_col, last_row, last_col) => {
                    match worksheet.autofilter(first_row, first_col, last_row, last_col) {
                        Ok(ws) => ws,
                        Err(e) => return Err(e.to_string()),
                    }
                }
                Sheet::MergeRange(first_row, first_col, last_row, last_col, data) => {
                    match merge_range(worksheet, first_row, first_col, last_row, last_col, data) {
                        Ok(ws) => ws,
                        Err(e) => return Err(e.to_string()),
                    }
                }
                Sheet::InsertNote(row, col, text, options) => {
                    match insert_note(worksheet, row, col, text, options) {
                        Ok(ws) => ws,
                        Err(e) => return Err(e.to_string()),
                    }
                }
                Sheet::SetTabColor(color_hex) => {
                    if let Some(color) = parse_hex_color(&color_hex) {
                        worksheet.set_tab_color(color)
                    } else {
                        worksheet
                    }
                }
                Sheet::SetAutofit => worksheet.autofit(),
                Sheet::Write(row, col, data) => match write_data(worksheet, row, col, data) {
                    Ok(ws) => ws,
                    Err(e) => return Err(e.to_string()),
                },
            };
        }
    }

    match workbook.save_to_buffer() {
        Ok(buf) => Ok(buf),
        Err(e) => Err(e.to_string()),
    }
}

fn set_workbook_properties(workbook: &mut Workbook, props: WorkbookProperties) {
    let mut doc_props = DocProperties::new();
    if let Some(author) = props.author {
        doc_props = doc_props.set_author(&author);
    }
    if let Some(title) = props.title {
        doc_props = doc_props.set_title(&title);
    }
    if let Some(subject) = props.subject {
        doc_props = doc_props.set_subject(&subject);
    }
    if let Some(manager) = props.manager {
        doc_props = doc_props.set_manager(&manager);
    }
    if let Some(company) = props.company {
        doc_props = doc_props.set_company(&company);
    }
    if let Some(category) = props.category {
        doc_props = doc_props.set_category(&category);
    }
    if let Some(keywords) = props.keywords {
        doc_props = doc_props.set_keywords(&keywords);
    }
    if let Some(comment) = props.comment {
        doc_props = doc_props.set_comment(&comment);
    }
    if let Some(status) = props.status {
        doc_props = doc_props.set_status(&status);
    }
    workbook.set_properties(&doc_props);
}

fn insert_note(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    text: String,
    options: NoteOptions,
) -> Result<&mut Worksheet, XlsxError> {
    let mut note = Note::new(&text);

    if let Some(author) = options.author {
        note = note.set_author(&author);
    }

    if let Some(visible) = options.visible {
        note = note.set_visible(visible);
    }

    if let Some(width) = options.width {
        note = note.set_width(width);
    }

    if let Some(height) = options.height {
        note = note.set_height(height);
    }

    worksheet.insert_note(row, col, &note)
}

fn merge_range<'a, 'b>(
    worksheet: &'a mut Worksheet,
    first_row: u32,
    first_col: u16,
    last_row: u32,
    last_col: u16,
    data: CellData<'b>,
) -> Result<&'a mut Worksheet, XlsxError> {
    match data {
        CellData::String(val) => worksheet.merge_range(
            first_row,
            first_col,
            last_row,
            last_col,
            &val,
            &Format::new(),
        ),
        CellData::StringWithFormat(val, formats) => {
            let format = apply_formats(Format::new(), &formats);
            worksheet.merge_range(first_row, first_col, last_row, last_col, &val, &format)
        }
        CellData::NumberWithFormat(val, formats) => {
            let format = apply_formats(Format::new(), &formats);
            // Write value to first cell, then merge the range with the same format
            worksheet.write_number_with_format(first_row, first_col, val, &format)?;
            worksheet.merge_range(first_row, first_col, last_row, last_col, "", &format)
        }
        CellData::Float(val) => {
            // Write number to first cell, then merge
            worksheet.write_number(first_row, first_col, val)?;
            worksheet.merge_range(first_row, first_col, last_row, last_col, "", &Format::new())
        }
        CellData::Boolean(val) => {
            // Write boolean to first cell, then merge
            worksheet.write_boolean(first_row, first_col, val)?;
            worksheet.merge_range(first_row, first_col, last_row, last_col, "", &Format::new())
        }
        CellData::BooleanWithFormat(val, formats) => {
            let format = apply_formats(Format::new(), &formats);
            worksheet.write_boolean_with_format(first_row, first_col, val, &format)?;
            worksheet.merge_range(first_row, first_col, last_row, last_col, "", &format)
        }
        CellData::Blank(formats) => {
            let format = apply_formats(Format::new(), &formats);
            worksheet.merge_range(first_row, first_col, last_row, last_col, "", &format)
        }
        CellData::DateWithFormat(iso8601, user_formats) => {
            let date_format =
                apply_formats(Format::new().set_num_format("yyyy-mm-dd"), &user_formats);

            let date = ExcelDateTime::parse_from_str(&iso8601)?;
            worksheet.write_with_format(first_row, first_col, &date, &date_format)?;
            worksheet.merge_range(first_row, first_col, last_row, last_col, "", &date_format)
        }
        CellData::DateTimeWithFormat(iso8601, user_formats) => {
            let date_format = apply_formats(
                Format::new().set_num_format("yyyy-mm-ddThh:mm:ss"),
                &user_formats,
            );

            let date = ExcelDateTime::parse_from_str(&iso8601)?;
            worksheet.write_with_format(first_row, first_col, &date, &date_format)?;
            worksheet.merge_range(first_row, first_col, last_row, last_col, "", &date_format)
        }
        // For other types that don't support merge_range, write to first cell only
        _ => write_data(worksheet, first_row, first_col, data),
    }
}

fn write_data<'a, 'b>(
    worksheet: &'a mut Worksheet,
    row: u32,
    col: u16,
    data: CellData<'b>,
) -> Result<&'a mut Worksheet, XlsxError> {
    match data {
        CellData::String(val) => worksheet.write(row, col, val),
        CellData::StringWithFormat(val, formats) => {
            let format = apply_formats(Format::new(), &formats);
            worksheet.write_with_format(row, col, val, &format)
        }
        CellData::NumberWithFormat(val, formats) => {
            let format = apply_formats(Format::new(), &formats);
            worksheet.write_number_with_format(row, col, val, &format)
        }

        CellData::Float(val) => worksheet.write(row, col, val),
        CellData::Date(iso8601) => {
            let date_format = Format::new().set_num_format("yyyy-mm-dd");
            let date = ExcelDateTime::parse_from_str(&iso8601)?;

            worksheet.write_with_format(row, col, &date, &date_format)
        }
        CellData::DateWithFormat(iso8601, user_formats) => {
            let date_format =
                apply_formats(Format::new().set_num_format("yyyy-mm-dd"), &user_formats);
            let date = ExcelDateTime::parse_from_str(&iso8601)?;

            worksheet.write_with_format(row, col, &date, &date_format)
        }
        CellData::DateTime(iso8601) => {
            let date_format = Format::new().set_num_format("yyyy-mm-ddThh:mm:ss");
            let date = ExcelDateTime::parse_from_str(&iso8601)?;

            worksheet.write_with_format(row, col, &date, &date_format)
        }
        CellData::DateTimeWithFormat(iso8601, user_formats) => {
            let date_format = apply_formats(
                Format::new().set_num_format("yyyy-mm-ddThh:mm:ss"),
                &user_formats,
            );
            let date = ExcelDateTime::parse_from_str(&iso8601)?;

            worksheet.write_with_format(row, col, &date, &date_format)
        }
        CellData::Formula(val) => worksheet.write(row, col, Formula::new(val)),
        CellData::FormulaWithFormat(val, formats) => {
            let format = apply_formats(Format::new(), &formats);
            worksheet.write_with_format(row, col, Formula::new(val), &format)
        }
        CellData::Boolean(val) => worksheet.write_boolean(row, col, val),
        CellData::BooleanWithFormat(val, formats) => {
            let format = apply_formats(Format::new(), &formats);
            worksheet.write_boolean_with_format(row, col, val, &format)
        }
        CellData::Url(url) => {
            let url_obj = Url::new(&url);
            worksheet.write_url(row, col, &url_obj)
        }
        CellData::UrlWithText(url, text) => {
            let url_obj = Url::new(&url);
            worksheet.write_url_with_text(row, col, &url_obj, &text)
        }
        CellData::UrlWithFormat(url, formats) => {
            let format = apply_formats(Format::new(), &formats);
            let url_obj = Url::new(&url);
            worksheet.write_url_with_format(row, col, &url_obj, &format)
        }
        CellData::UrlWithTextAndFormat(url, text, formats) => {
            let format = apply_formats(Format::new(), &formats);
            let url_obj = Url::new(&url).set_text(&text);
            worksheet.write_url_with_format(row, col, &url_obj, &format)
        }
        CellData::Blank(formats) => {
            let format = apply_formats(Format::new(), &formats);
            worksheet.write_blank(row, col, &format)
        }
        CellData::ImagePath(val) => {
            let image = Image::new(val)?;

            worksheet.insert_image(row, col, &image)
        }
        CellData::Image(binary) => {
            let image = Image::new_from_buffer(binary.as_slice())?;

            worksheet.insert_image(row, col, &image)
        }
        CellData::RichString(segments) => {
            write_rich_string_helper(worksheet, row, col, &segments, None)
        }
        CellData::RichStringWithFormat(segments, cell_formats) => {
            let cell_format = apply_formats(Format::new(), &cell_formats);
            write_rich_string_helper(worksheet, row, col, &segments, Some(cell_format))
        }
    }
}

fn write_rich_string_helper<'a>(
    worksheet: &'a mut Worksheet,
    row: u32,
    col: u16,
    segments: &[(String, Vec<CellFormat>)],
    cell_format: Option<Format>,
) -> Result<&'a mut Worksheet, XlsxError> {
    // Build format objects for each segment
    let segment_formats: Vec<Format> = segments
        .iter()
        .map(|(_, formats)| apply_formats(Format::new(), formats))
        .collect();

    // Build the segments array with references
    let rich_segments: Vec<(&Format, &str)> = segments
        .iter()
        .zip(segment_formats.iter())
        .map(|((text, _), format)| (format, text.as_str()))
        .collect();

    match cell_format {
        Some(format) => worksheet.write_rich_string_with_format(row, col, &rich_segments, &format),
        None => worksheet.write_rich_string(row, col, &rich_segments),
    }
}

fn apply_formats(mut format: Format, formats: &[CellFormat]) -> Format {
    for fmt in formats {
        format = match fmt {
            CellFormat::Bold => format.set_bold(),
            CellFormat::NumFormat(format_string) => format.set_num_format(format_string),
            CellFormat::Align(pos) => match pos {
                CellAlignPos::Center => format.set_align(FormatAlign::Center),
                CellAlignPos::Right => format.set_align(FormatAlign::Right),
                CellAlignPos::Left => format.set_align(FormatAlign::Left),
            },
            CellFormat::BgColor(color_hex) => {
                if let Some(color) = parse_hex_color(color_hex) {
                    format.set_background_color(color)
                } else {
                    format
                }
            }
            CellFormat::Pattern(pattern) => match pattern {
                CellPattern::Solid => format.set_pattern(FormatPattern::Solid),
                CellPattern::None => format.set_pattern(FormatPattern::None),
                CellPattern::Gray125 => format.set_pattern(FormatPattern::Gray125),
                CellPattern::Gray0625 => format.set_pattern(FormatPattern::Gray0625),
            },
            CellFormat::FontColor(color_hex) => {
                if let Some(color) = parse_hex_color(color_hex) {
                    format.set_font_color(color)
                } else {
                    format
                }
            }
            CellFormat::Italic => format.set_italic(),
            CellFormat::Underline(style) => match style {
                UnderlineStyle::Single => format.set_underline(FormatUnderline::Single),
                UnderlineStyle::Double => format.set_underline(FormatUnderline::Double),
                UnderlineStyle::SingleAccounting => {
                    format.set_underline(FormatUnderline::SingleAccounting)
                }
                UnderlineStyle::DoubleAccounting => {
                    format.set_underline(FormatUnderline::DoubleAccounting)
                }
            },
            CellFormat::Strikethrough => format.set_font_strikethrough(),
            CellFormat::FontSize(size) => format.set_font_size(*size),
            CellFormat::FontName(name) => format.set_font_name(name),
            CellFormat::Superscript => format.set_font_script(FormatScript::Superscript),
            CellFormat::Subscript => format.set_font_script(FormatScript::Subscript),
            CellFormat::Border(style) => {
                let border_style = convert_border_style(style);
                format.set_border(border_style)
            }
            CellFormat::BorderTop(style) => {
                let border_style = convert_border_style(style);
                format.set_border_top(border_style)
            }
            CellFormat::BorderBottom(style) => {
                let border_style = convert_border_style(style);
                format.set_border_bottom(border_style)
            }
            CellFormat::BorderLeft(style) => {
                let border_style = convert_border_style(style);
                format.set_border_left(border_style)
            }
            CellFormat::BorderRight(style) => {
                let border_style = convert_border_style(style);
                format.set_border_right(border_style)
            }
            CellFormat::BorderColor(color_hex) => {
                if let Some(color) = parse_hex_color(color_hex) {
                    format.set_border_color(color)
                } else {
                    format
                }
            }
            CellFormat::BorderTopColor(color_hex) => {
                if let Some(color) = parse_hex_color(color_hex) {
                    format.set_border_top_color(color)
                } else {
                    format
                }
            }
            CellFormat::BorderBottomColor(color_hex) => {
                if let Some(color) = parse_hex_color(color_hex) {
                    format.set_border_bottom_color(color)
                } else {
                    format
                }
            }
            CellFormat::BorderLeftColor(color_hex) => {
                if let Some(color) = parse_hex_color(color_hex) {
                    format.set_border_left_color(color)
                } else {
                    format
                }
            }
            CellFormat::BorderRightColor(color_hex) => {
                if let Some(color) = parse_hex_color(color_hex) {
                    format.set_border_right_color(color)
                } else {
                    format
                }
            }
            CellFormat::TextWrap => format.set_text_wrap(),
            CellFormat::Valign(pos) => match pos {
                CellVAlignPos::Top => format.set_align(FormatAlign::Top),
                CellVAlignPos::Center => format.set_align(FormatAlign::VerticalCenter),
                CellVAlignPos::Bottom => format.set_align(FormatAlign::Bottom),
                CellVAlignPos::Justify => format.set_align(FormatAlign::VerticalJustify),
                CellVAlignPos::Distributed => format.set_align(FormatAlign::VerticalDistributed),
            },
            CellFormat::Rotation(angle) => format.set_rotation(*angle),
            CellFormat::Shrink => format.set_shrink(),
            CellFormat::Indent(level) => format.set_indent(*level),
        };
    }

    format
}

/// Parses a hex color string (e.g., "#FF0000" or "FF0000") into a Color.
/// Returns None if the hex string is invalid.
fn parse_hex_color(color_hex: &str) -> Option<Color> {
    let hex_str = color_hex.trim_start_matches('#');
    u32::from_str_radix(hex_str, 16).ok().map(Color::from)
}

fn convert_border_style(style: &BorderStyle) -> FormatBorder {
    match style {
        BorderStyle::Thin => FormatBorder::Thin,
        BorderStyle::Medium => FormatBorder::Medium,
        BorderStyle::Thick => FormatBorder::Thick,
        BorderStyle::Dashed => FormatBorder::Dashed,
        BorderStyle::Dotted => FormatBorder::Dotted,
        BorderStyle::Double => FormatBorder::Double,
        BorderStyle::Hair => FormatBorder::Hair,
        BorderStyle::MediumDashed => FormatBorder::MediumDashed,
        BorderStyle::DashDot => FormatBorder::DashDot,
        BorderStyle::MediumDashDot => FormatBorder::MediumDashDot,
        BorderStyle::DashDotDot => FormatBorder::DashDotDot,
        BorderStyle::MediumDashDotDot => FormatBorder::MediumDashDotDot,
        BorderStyle::SlantDashDot => FormatBorder::SlantDashDot,
    }
}

rustler::init!("Elixir.XlsxWriter.RustXlsxWriter");
