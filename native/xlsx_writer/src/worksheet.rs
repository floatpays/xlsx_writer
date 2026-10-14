//! Executing a sheet's instructions on a rust_xlsxwriter `Worksheet`.

use rust_xlsxwriter::{ExcelDateTime, Format, Formula, Image, Note, Url, Worksheet};

use crate::error::{WriteError, WriteResult};
use crate::format::{apply_formats, cell_format, parse_hex_color};
use crate::protocol::{column_number, row_number, CellData, CellFormat, Instruction, NoteOptions};

/// Excel's maximum column width, in character units.
const MAX_COLUMN_WIDTH: f64 = 255.0;
/// Excel's maximum row height, in points.
const MAX_ROW_HEIGHT: f64 = 409.0;
/// Excel's maximum length of a comment author name.
const MAX_AUTHOR_LENGTH: usize = 52;

/// Per-sheet state that has to be applied after all instructions.
#[derive(Default)]
pub struct SheetState {
    autofit_requested: bool,
    /// Explicit column widths as `(first_column, last_column, width)`, in
    /// the order they were set, re-applied after autofit so they win.
    column_widths: Vec<(u16, u16, f64)>,
}

pub fn apply_instruction(
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
pub fn finish_sheet(worksheet: &mut Worksheet, state: SheetState) -> WriteResult<()> {
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn checks_sizes() {
        assert_eq!(column_width(25.5), Ok(25.5));
        assert!(column_width(-1.0).is_err());
        assert!(column_width(256.0).is_err());
        assert!(column_width(f64::NAN).is_err());

        assert_eq!(row_height(30.0), Ok(30.0));
        assert!(row_height(409.5).is_err());

        assert_eq!(note_size(100.4, "width"), Ok(100));
        assert!(note_size(0.2, "width").is_err());
        assert!(note_size(f64::INFINITY, "width").is_err());
    }
}
