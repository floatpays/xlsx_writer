use std::fmt;

use rust_xlsxwriter::{
    Color, DocProperties, ExcelDateTime, Format, FormatAlign, FormatBorder, FormatPattern,
    FormatScript, FormatUnderline, Formula, Image, Note, Url, Workbook, Worksheet, XlsxError,
};
use rustler::{Binary, Env, NifStruct, NifTaggedEnum, OwnedBinary, Term};

/// Excel allows rows 0..=1_048_575.
const MAX_ROW: u32 = 1_048_575;
/// Excel allows columns 0..=16_383.
const MAX_COLUMN: u16 = 16_383;
/// Excel's maximum column width, in character units.
const MAX_COLUMN_WIDTH: f64 = 255.0;
/// Excel's maximum row height, in points.
const MAX_ROW_HEIGHT: f64 = 409.0;
/// Excel's maximum font size, in points.
const MAX_FONT_SIZE: f64 = 409.0;
/// Excel's maximum indent level.
const MAX_INDENT: u8 = 250;
/// Excel's maximum length of a comment author name.
const MAX_AUTHOR_LENGTH: usize = 52;

const DEFAULT_DATE_FORMAT: &str = "yyyy-mm-dd";
const DEFAULT_DATE_TIME_FORMAT: &str = "yyyy-mm-ddThh:mm:ss";

/// Error returned to Elixir as the `reason` in `{:error, reason}`.
#[derive(Debug, PartialEq)]
struct WriteError(String);

impl WriteError {
    fn new(message: impl Into<String>) -> Self {
        WriteError(message.into())
    }
}

impl From<XlsxError> for WriteError {
    fn from(error: XlsxError) -> Self {
        WriteError(error.to_string())
    }
}

impl fmt::Display for WriteError {
    fn fmt(&self, formatter: &mut fmt::Formatter) -> fmt::Result {
        formatter.write_str(&self.0)
    }
}

type WriteResult<T> = Result<T, WriteError>;

#[derive(NifTaggedEnum, PartialEq, Debug)]
enum CellAlignPos {
    Center,
    Left,
    Right,
}

#[derive(NifTaggedEnum, PartialEq, Debug)]
enum CellVAlignPos {
    Top,
    Center,
    Bottom,
    Justify,
    Distributed,
}

#[derive(NifTaggedEnum, PartialEq, Debug)]
enum CellPattern {
    Solid,
    None,
    Gray125,
    Gray0625,
}

#[derive(NifTaggedEnum, PartialEq, Debug)]
enum UnderlineStyle {
    Single,
    Double,
    SingleAccounting,
    DoubleAccounting,
}

#[derive(NifTaggedEnum, PartialEq, Debug)]
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

#[derive(NifTaggedEnum, PartialEq, Debug)]
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
    FontSize(f64),
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
    Rotation(i64),
    Shrink,
    Indent(i64),
}

#[derive(NifStruct)]
#[module = "XlsxWriter.NoteOptions"]
struct NoteOptions {
    author: Option<String>,
    visible: Option<bool>,
    width: Option<f64>,
    height: Option<f64>,
}

/// A cell value plus its format list. An empty format list means "no
/// format", which matters for URLs (they keep Excel's hyperlink style).
#[derive(NifTaggedEnum)]
enum CellData<'a> {
    String(String, Vec<CellFormat>),
    Number(f64, Vec<CellFormat>),
    Boolean(bool, Vec<CellFormat>),
    Date(String, Vec<CellFormat>),
    DateTime(String, Vec<CellFormat>),
    Formula(String, Vec<CellFormat>),
    Url(String, Option<String>, Vec<CellFormat>),
    Blank(Vec<CellFormat>),
    RichString(Vec<(String, Vec<CellFormat>)>, Vec<CellFormat>),
    Image(Binary<'a>),
    ImagePath(String),
}

/// Positions are decoded as `i64` and sizes as `f64` so that out-of-range or
/// fractional values reach our own checks instead of being truncated by the
/// decoder.
#[derive(NifTaggedEnum)]
enum Instruction<'a> {
    Write(i64, i64, CellData<'a>),
    MergeRange(i64, i64, i64, i64, CellData<'a>),
    SetColumnWidth(i64, f64),
    SetRowHeight(i64, f64),
    SetColumnRangeWidth(i64, i64, f64),
    SetRowRangeHeight(i64, i64, f64),
    SetFreezePanes(i64, i64),
    SetRowHidden(i64),
    SetColumnHidden(i64),
    SetAutofilter(i64, i64, i64, i64),
    InsertNote(i64, i64, String, NoteOptions),
    SetTabColor(String),
    SetAutofit,
}

#[derive(NifStruct)]
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

/// Per-sheet state that has to be applied after all instructions.
#[derive(Default)]
struct SheetState {
    autofit_requested: bool,
    /// Explicit column widths as `(first_column, last_column, width)`, in
    /// the order they were set, re-applied after autofit so they win.
    column_widths: Vec<(u16, u16, f64)>,
}

#[rustler::nif(schedule = "DirtyCpu")]
fn write<'a>(
    env: Env<'a>,
    sheets: Vec<Term<'a>>,
    properties: Term<'a>,
) -> Result<Binary<'a>, String> {
    let buffer = write_workbook(sheets, properties).map_err(|error| error.0)?;

    let mut binary = OwnedBinary::new(buffer.len())
        .ok_or_else(|| "could not allocate a binary for the workbook".to_string())?;
    binary.as_mut_slice().copy_from_slice(&buffer);

    Ok(binary.release(env))
}

fn write_workbook(sheets: Vec<Term>, properties: Term) -> WriteResult<Vec<u8>> {
    let mut workbook = Workbook::new();

    let properties: Option<WorkbookProperties> = properties
        .decode()
        .map_err(|_| WriteError::new(format!("invalid workbook properties: {properties:?}")))?;

    if let Some(properties) = properties {
        workbook.set_properties(&document_properties(properties));
    }

    for (sheet_index, sheet) in sheets.iter().enumerate() {
        let (sheet_name, instructions): (String, Vec<Term>) = sheet.decode().map_err(|_| {
            WriteError::new(format!("invalid sheet #{}: {sheet:?}", sheet_index + 1))
        })?;

        let worksheet = workbook.add_worksheet();
        worksheet.set_name(&sheet_name)?;

        let mut state = SheetState::default();

        for (instruction_index, term) in instructions.iter().enumerate() {
            let instruction: Instruction = term.decode().map_err(|_| {
                WriteError::new(format!(
                    "invalid instruction #{} in sheet {sheet_name:?}: {term:?}",
                    instruction_index + 1
                ))
            })?;

            apply_instruction(worksheet, &mut state, instruction)?;
        }

        finish_sheet(worksheet, state)?;
    }

    Ok(workbook.save_to_buffer()?)
}

fn document_properties(properties: WorkbookProperties) -> DocProperties {
    let mut document = DocProperties::new();

    if let Some(author) = properties.author {
        document = document.set_author(&author);
    }
    if let Some(title) = properties.title {
        document = document.set_title(&title);
    }
    if let Some(subject) = properties.subject {
        document = document.set_subject(&subject);
    }
    if let Some(manager) = properties.manager {
        document = document.set_manager(&manager);
    }
    if let Some(company) = properties.company {
        document = document.set_company(&company);
    }
    if let Some(category) = properties.category {
        document = document.set_category(&category);
    }
    if let Some(keywords) = properties.keywords {
        document = document.set_keywords(&keywords);
    }
    if let Some(comment) = properties.comment {
        document = document.set_comment(&comment);
    }
    if let Some(status) = properties.status {
        document = document.set_status(&status);
    }

    document
}

fn apply_instruction(
    worksheet: &mut Worksheet,
    state: &mut SheetState,
    instruction: Instruction,
) -> WriteResult<()> {
    match instruction {
        Instruction::Write(row, column, data) => {
            let format = cell_format(&data)?;
            write_cell(
                worksheet,
                row_number(row)?,
                column_number(column)?,
                &data,
                format.as_ref(),
            )?;
        }
        Instruction::MergeRange(first_row, first_column, last_row, last_column, data) => {
            merge_range(
                worksheet,
                row_number(first_row)?,
                column_number(first_column)?,
                row_number(last_row)?,
                column_number(last_column)?,
                &data,
            )?;
        }
        Instruction::SetColumnWidth(column, width) => {
            let column = column_number(column)?;
            let width = column_width(width)?;
            worksheet.set_column_width(column, width)?;
            state.column_widths.push((column, column, width));
        }
        Instruction::SetColumnRangeWidth(first_column, last_column, width) => {
            let first_column = column_number(first_column)?;
            let last_column = column_number(last_column)?;
            let width = column_width(width)?;
            worksheet.set_column_range_width(first_column, last_column, width)?;
            state.column_widths.push((first_column, last_column, width));
        }
        Instruction::SetRowHeight(row, height) => {
            worksheet.set_row_height(row_number(row)?, row_height(height)?)?;
        }
        Instruction::SetRowRangeHeight(first_row, last_row, height) => {
            let first_row = row_number(first_row)?;
            let last_row = row_number(last_row)?;
            let height = row_height(height)?;

            if first_row > last_row {
                return Err(WriteError::new(format!(
                    "first row {first_row} is after last row {last_row}"
                )));
            }

            for row in first_row..=last_row {
                worksheet.set_row_height(row, height)?;
            }
        }
        Instruction::SetFreezePanes(row, column) => {
            worksheet.set_freeze_panes(row_number(row)?, column_number(column)?)?;
        }
        Instruction::SetRowHidden(row) => {
            worksheet.set_row_hidden(row_number(row)?)?;
        }
        Instruction::SetColumnHidden(column) => {
            worksheet.set_column_hidden(column_number(column)?)?;
        }
        Instruction::SetAutofilter(first_row, first_column, last_row, last_column) => {
            worksheet.autofilter(
                row_number(first_row)?,
                column_number(first_column)?,
                row_number(last_row)?,
                column_number(last_column)?,
            )?;
        }
        Instruction::InsertNote(row, column, text, options) => {
            let note = build_note(&text, options)?;
            worksheet.insert_note(row_number(row)?, column_number(column)?, &note)?;
        }
        Instruction::SetTabColor(color) => {
            worksheet.set_tab_color(parse_hex_color(&color)?);
        }
        Instruction::SetAutofit => state.autofit_requested = true,
    }

    Ok(())
}

/// Autofit runs once, after every cell is written, and explicit column widths
/// are re-applied afterwards so that they take precedence over it.
fn finish_sheet(worksheet: &mut Worksheet, state: SheetState) -> WriteResult<()> {
    if state.autofit_requested {
        worksheet.autofit();

        for (first_column, last_column, width) in state.column_widths {
            worksheet.set_column_range_width(first_column, last_column, width)?;
        }
    }

    Ok(())
}

fn build_note(text: &str, options: NoteOptions) -> WriteResult<Note> {
    let mut note = Note::new(text);

    if let Some(author) = options.author {
        if author.chars().count() > MAX_AUTHOR_LENGTH {
            return Err(WriteError::new(format!(
                "comment author must be at most {MAX_AUTHOR_LENGTH} characters, got {author:?}"
            )));
        }
        note = note.set_author(author);
    }

    if let Some(visible) = options.visible {
        note = note.set_visible(visible);
    }

    if let Some(width) = options.width {
        note = note.set_width(note_size(width, "width")?);
    }

    if let Some(height) = options.height {
        note = note.set_height(note_size(height, "height")?);
    }

    Ok(note)
}

/// Merges the range with the cell's format, then writes the typed value into
/// the first cell with the same format. `Worksheet::merge_range` only takes a
/// string, so writing the value afterwards keeps numbers, booleans, dates and
/// formulas typed.
fn merge_range(
    worksheet: &mut Worksheet,
    first_row: u32,
    first_column: u16,
    last_row: u32,
    last_column: u16,
    data: &CellData,
) -> WriteResult<()> {
    if matches!(data, CellData::Image(_) | CellData::ImagePath(_)) {
        return Err(WriteError::new(
            "images cannot be written to a merged range",
        ));
    }

    let format = cell_format(data)?;
    let merge_format = format.clone().unwrap_or_default();

    worksheet.merge_range(
        first_row,
        first_column,
        last_row,
        last_column,
        "",
        &merge_format,
    )?;

    write_cell(worksheet, first_row, first_column, data, format.as_ref())
}

/// The format a cell is written with: `None` when the cell has no format of
/// its own, so that rust_xlsxwriter applies its defaults (e.g. the hyperlink
/// style for URLs).
fn cell_format(data: &CellData) -> WriteResult<Option<Format>> {
    match data {
        CellData::Date(_, formats) => {
            apply_formats(Format::new().set_num_format(DEFAULT_DATE_FORMAT), formats).map(Some)
        }
        CellData::DateTime(_, formats) => apply_formats(
            Format::new().set_num_format(DEFAULT_DATE_TIME_FORMAT),
            formats,
        )
        .map(Some),
        CellData::String(_, formats)
        | CellData::Number(_, formats)
        | CellData::Boolean(_, formats)
        | CellData::Formula(_, formats)
        | CellData::Url(_, _, formats)
        | CellData::Blank(formats)
        | CellData::RichString(_, formats) => optional_format(formats),
        CellData::Image(_) | CellData::ImagePath(_) => Ok(None),
    }
}

fn optional_format(formats: &[CellFormat]) -> WriteResult<Option<Format>> {
    if formats.is_empty() {
        Ok(None)
    } else {
        apply_formats(Format::new(), formats).map(Some)
    }
}

fn write_cell(
    worksheet: &mut Worksheet,
    row: u32,
    column: u16,
    data: &CellData,
    format: Option<&Format>,
) -> WriteResult<()> {
    match (data, format) {
        (CellData::String(text, _), Some(format)) => {
            worksheet.write_string_with_format(row, column, text, format)?;
        }
        (CellData::String(text, _), None) => {
            worksheet.write_string(row, column, text)?;
        }
        (CellData::Number(number, _), Some(format)) => {
            worksheet.write_number_with_format(row, column, *number, format)?;
        }
        (CellData::Number(number, _), None) => {
            worksheet.write_number(row, column, *number)?;
        }
        (CellData::Boolean(boolean, _), Some(format)) => {
            worksheet.write_boolean_with_format(row, column, *boolean, format)?;
        }
        (CellData::Boolean(boolean, _), None) => {
            worksheet.write_boolean(row, column, *boolean)?;
        }
        (CellData::Date(iso8601, _) | CellData::DateTime(iso8601, _), format) => {
            let date_time = ExcelDateTime::parse_from_str(iso8601)?;
            let format = format.cloned().unwrap_or_default();
            worksheet.write_datetime_with_format(row, column, &date_time, &format)?;
        }
        (CellData::Formula(formula, _), Some(format)) => {
            worksheet.write_formula_with_format(row, column, Formula::new(formula), format)?;
        }
        (CellData::Formula(formula, _), None) => {
            worksheet.write_formula(row, column, Formula::new(formula))?;
        }
        (CellData::Url(url, text, _), format) => {
            let mut link = Url::new(url);
            if let Some(text) = text {
                link = link.set_text(text);
            }
            match format {
                Some(format) => worksheet.write_url_with_format(row, column, link, format)?,
                None => worksheet.write_url(row, column, link)?,
            };
        }
        (CellData::Blank(_), format) => {
            worksheet.write_blank(row, column, &format.cloned().unwrap_or_default())?;
        }
        (CellData::RichString(segments, _), format) => {
            write_rich_string(worksheet, row, column, segments, format)?;
        }
        (CellData::Image(binary), _) => {
            let image = Image::new_from_buffer(binary.as_slice())?;
            worksheet.insert_image(row, column, &image)?;
        }
        (CellData::ImagePath(path), _) => {
            let image = Image::new(path)?;
            worksheet.insert_image(row, column, &image)?;
        }
    }

    Ok(())
}

fn write_rich_string(
    worksheet: &mut Worksheet,
    row: u32,
    column: u16,
    segments: &[(String, Vec<CellFormat>)],
    cell_format: Option<&Format>,
) -> WriteResult<()> {
    let segment_formats = segments
        .iter()
        .map(|(_, formats)| apply_formats(Format::new(), formats))
        .collect::<WriteResult<Vec<Format>>>()?;

    let rich_segments: Vec<(&Format, &str)> = segments
        .iter()
        .zip(segment_formats.iter())
        .map(|((text, _), format)| (format, text.as_str()))
        .collect();

    match cell_format {
        Some(format) => {
            worksheet.write_rich_string_with_format(row, column, &rich_segments, format)?
        }
        None => worksheet.write_rich_string(row, column, &rich_segments)?,
    };

    Ok(())
}

fn apply_formats(mut format: Format, formats: &[CellFormat]) -> WriteResult<Format> {
    for cell_format in formats {
        format = match cell_format {
            CellFormat::Bold => format.set_bold(),
            CellFormat::NumFormat(format_string) => format.set_num_format(format_string),
            CellFormat::Align(position) => match position {
                CellAlignPos::Center => format.set_align(FormatAlign::Center),
                CellAlignPos::Right => format.set_align(FormatAlign::Right),
                CellAlignPos::Left => format.set_align(FormatAlign::Left),
            },
            CellFormat::BgColor(color) => format.set_background_color(parse_hex_color(color)?),
            CellFormat::Pattern(pattern) => match pattern {
                CellPattern::Solid => format.set_pattern(FormatPattern::Solid),
                CellPattern::None => format.set_pattern(FormatPattern::None),
                CellPattern::Gray125 => format.set_pattern(FormatPattern::Gray125),
                CellPattern::Gray0625 => format.set_pattern(FormatPattern::Gray0625),
            },
            CellFormat::FontColor(color) => format.set_font_color(parse_hex_color(color)?),
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
            CellFormat::FontSize(size) => format.set_font_size(font_size(*size)?),
            CellFormat::FontName(name) => format.set_font_name(name),
            CellFormat::Superscript => format.set_font_script(FormatScript::Superscript),
            CellFormat::Subscript => format.set_font_script(FormatScript::Subscript),
            CellFormat::Border(style) => format.set_border(border_style(style)),
            CellFormat::BorderTop(style) => format.set_border_top(border_style(style)),
            CellFormat::BorderBottom(style) => format.set_border_bottom(border_style(style)),
            CellFormat::BorderLeft(style) => format.set_border_left(border_style(style)),
            CellFormat::BorderRight(style) => format.set_border_right(border_style(style)),
            CellFormat::BorderColor(color) => format.set_border_color(parse_hex_color(color)?),
            CellFormat::BorderTopColor(color) => {
                format.set_border_top_color(parse_hex_color(color)?)
            }
            CellFormat::BorderBottomColor(color) => {
                format.set_border_bottom_color(parse_hex_color(color)?)
            }
            CellFormat::BorderLeftColor(color) => {
                format.set_border_left_color(parse_hex_color(color)?)
            }
            CellFormat::BorderRightColor(color) => {
                format.set_border_right_color(parse_hex_color(color)?)
            }
            CellFormat::TextWrap => format.set_text_wrap(),
            CellFormat::Valign(position) => match position {
                CellVAlignPos::Top => format.set_align(FormatAlign::Top),
                CellVAlignPos::Center => format.set_align(FormatAlign::VerticalCenter),
                CellVAlignPos::Bottom => format.set_align(FormatAlign::Bottom),
                CellVAlignPos::Justify => format.set_align(FormatAlign::VerticalJustify),
                CellVAlignPos::Distributed => format.set_align(FormatAlign::VerticalDistributed),
            },
            CellFormat::Rotation(angle) => format.set_rotation(rotation(*angle)?),
            CellFormat::Shrink => format.set_shrink(),
            CellFormat::Indent(level) => format.set_indent(indent(*level)?),
        };
    }

    Ok(format)
}

/// Parses `"#RRGGBB"` or `"RRGGBB"` (case-insensitive). Anything else is an
/// error rather than being ignored or misread.
fn parse_hex_color(color: &str) -> WriteResult<Color> {
    let hex = color.strip_prefix('#').unwrap_or(color);

    if hex.len() != 6 || !hex.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(WriteError::new(format!(
            "invalid color {color:?}, expected a hex string like \"#RRGGBB\""
        )));
    }

    u32::from_str_radix(hex, 16)
        .map(Color::RGB)
        .map_err(|_| WriteError::new(format!("invalid color {color:?}")))
}

fn row_number(row: i64) -> WriteResult<u32> {
    u32::try_from(row)
        .ok()
        .filter(|row| *row <= MAX_ROW)
        .ok_or_else(|| WriteError::new(format!("row {row} is outside Excel's range 0..{MAX_ROW}")))
}

fn column_number(column: i64) -> WriteResult<u16> {
    u16::try_from(column)
        .ok()
        .filter(|column| *column <= MAX_COLUMN)
        .ok_or_else(|| {
            WriteError::new(format!(
                "column {column} is outside Excel's range 0..{MAX_COLUMN}"
            ))
        })
}

fn bounded_size(value: f64, maximum: f64, name: &str) -> WriteResult<f64> {
    if value.is_finite() && (0.0..=maximum).contains(&value) {
        Ok(value)
    } else {
        Err(WriteError::new(format!(
            "{name} {value} is outside Excel's range 0..{maximum}"
        )))
    }
}

fn column_width(width: f64) -> WriteResult<f64> {
    bounded_size(width, MAX_COLUMN_WIDTH, "column width")
}

fn row_height(height: f64) -> WriteResult<f64> {
    bounded_size(height, MAX_ROW_HEIGHT, "row height")
}

fn font_size(size: f64) -> WriteResult<f64> {
    if size > 0.0 {
        bounded_size(size, MAX_FONT_SIZE, "font size")
    } else {
        Err(WriteError::new(format!(
            "font size {size} must be greater than 0"
        )))
    }
}

fn rotation(angle: i64) -> WriteResult<i16> {
    match angle {
        -90..=90 | 270 => Ok(angle as i16),
        _ => Err(WriteError::new(format!(
            "rotation {angle} must be between -90 and 90, or 270"
        ))),
    }
}

fn indent(level: i64) -> WriteResult<u8> {
    u8::try_from(level)
        .ok()
        .filter(|level| *level <= MAX_INDENT)
        .ok_or_else(|| {
            WriteError::new(format!(
                "indent {level} is outside the range 0..{MAX_INDENT}"
            ))
        })
}

/// Comment box sizes are whole pixels of at least 1 (rust_xlsxwriter ignores 0).
fn note_size(value: f64, name: &str) -> WriteResult<u32> {
    let rounded = value.round();

    if value.is_finite() && rounded >= 1.0 && rounded <= f64::from(u32::MAX) {
        Ok(rounded as u32)
    } else {
        Err(WriteError::new(format!(
            "comment {name} {value} must be a number of pixels >= 1"
        )))
    }
}

fn border_style(style: &BorderStyle) -> FormatBorder {
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_six_digit_hex_colors() {
        assert_eq!(parse_hex_color("#FF0000"), Ok(Color::RGB(0xFF0000)));
        assert_eq!(parse_hex_color("#00ff7f"), Ok(Color::RGB(0x00FF7F)));
        assert_eq!(parse_hex_color("4472C4"), Ok(Color::RGB(0x4472C4)));
    }

    #[test]
    fn rejects_other_color_strings() {
        for color in [
            "#FFF",
            "FFF",
            "red",
            "#FFFFFFFF",
            "",
            "#",
            "#GGGGGG",
            "#+12345",
        ] {
            assert!(
                parse_hex_color(color).is_err(),
                "{color:?} should be rejected"
            );
        }
    }

    #[test]
    fn checks_row_and_column_bounds() {
        assert_eq!(row_number(0), Ok(0));
        assert_eq!(row_number(1_048_575), Ok(1_048_575));
        assert!(row_number(1_048_576).is_err());
        assert!(row_number(-1).is_err());
        assert!(row_number(i64::from(u32::MAX) + 1).is_err());

        assert_eq!(column_number(16_383), Ok(16_383));
        assert!(column_number(16_384).is_err());
        assert!(column_number(70_000).is_err());
        assert!(column_number(-1).is_err());
    }

    #[test]
    fn checks_sizes() {
        assert_eq!(column_width(25.5), Ok(25.5));
        assert!(column_width(-1.0).is_err());
        assert!(column_width(256.0).is_err());
        assert!(column_width(f64::NAN).is_err());

        assert_eq!(row_height(30.0), Ok(30.0));
        assert!(row_height(409.5).is_err());

        assert_eq!(font_size(11.5), Ok(11.5));
        assert!(font_size(0.0).is_err());
        assert!(font_size(410.0).is_err());

        assert_eq!(note_size(100.4, "width"), Ok(100));
        assert!(note_size(0.2, "width").is_err());
        assert!(note_size(f64::INFINITY, "width").is_err());
    }

    #[test]
    fn checks_rotation_and_indent() {
        assert_eq!(rotation(-90), Ok(-90));
        assert_eq!(rotation(90), Ok(90));
        assert_eq!(rotation(270), Ok(270));
        assert!(rotation(91).is_err());
        assert!(rotation(65_536 + 45).is_err());

        assert_eq!(indent(250), Ok(250));
        assert!(indent(251).is_err());
        assert!(indent(-1).is_err());
    }

    #[test]
    fn rejects_invalid_colors_in_formats() {
        assert!(apply_formats(Format::new(), &[CellFormat::BgColor("#FFF".into())]).is_err());
        assert!(apply_formats(Format::new(), &[CellFormat::FontColor("#FF0000".into())]).is_ok());
    }
}
