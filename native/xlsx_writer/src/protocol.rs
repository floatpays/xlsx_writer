//! The terms `XlsxWriter` sends to the NIF. They are private to this library:
//! `lib/xlsx_writer.ex` builds them and this module decodes them.

use rustler::{Binary, NifStruct, NifTaggedEnum};

use crate::error::{WriteError, WriteResult};

/// Excel allows rows 0..=1_048_575.
const MAX_ROW: u32 = 1_048_575;
/// Excel allows columns 0..=16_383.
const MAX_COLUMN: u16 = 16_383;

#[derive(NifTaggedEnum, PartialEq, Debug)]
pub enum CellAlignPos {
    Center,
    Left,
    Right,
}

#[derive(NifTaggedEnum, PartialEq, Debug)]
pub enum CellVAlignPos {
    Top,
    Center,
    Bottom,
    Justify,
    Distributed,
}

#[derive(NifTaggedEnum, PartialEq, Debug)]
pub enum CellPattern {
    Solid,
    None,
    Gray125,
    Gray0625,
}

#[derive(NifTaggedEnum, PartialEq, Debug)]
pub enum UnderlineStyle {
    Single,
    Double,
    SingleAccounting,
    DoubleAccounting,
}

#[derive(NifTaggedEnum, PartialEq, Debug)]
pub enum BorderStyle {
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
pub enum CellFormat {
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
pub struct NoteOptions {
    pub author: Option<String>,
    pub visible: Option<bool>,
    pub width: Option<f64>,
    pub height: Option<f64>,
}

/// A cell value plus its format list. An empty format list means "no
/// format", which matters for URLs (they keep Excel's hyperlink style).
#[derive(NifTaggedEnum)]
pub enum CellData<'a> {
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
}

/// Positions are decoded as `i64` and sizes as `f64` so that out-of-range or
/// fractional values reach our own checks instead of being truncated by the
/// decoder.
#[derive(NifTaggedEnum)]
pub enum Instruction<'a> {
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
pub struct WorkbookProperties {
    pub author: Option<String>,
    pub title: Option<String>,
    pub subject: Option<String>,
    pub manager: Option<String>,
    pub company: Option<String>,
    pub category: Option<String>,
    pub keywords: Option<String>,
    pub comment: Option<String>,
    pub status: Option<String>,
}

pub fn row_number(row: i64) -> WriteResult<u32> {
    u32::try_from(row)
        .ok()
        .filter(|row| *row <= MAX_ROW)
        .ok_or_else(|| WriteError::new(format!("row {row} is outside Excel's range 0..{MAX_ROW}")))
}

pub fn column_number(column: i64) -> WriteResult<u16> {
    u16::try_from(column)
        .ok()
        .filter(|column| *column <= MAX_COLUMN)
        .ok_or_else(|| {
            WriteError::new(format!(
                "column {column} is outside Excel's range 0..{MAX_COLUMN}"
            ))
        })
}

#[cfg(test)]
mod tests {
    use super::*;

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
}
