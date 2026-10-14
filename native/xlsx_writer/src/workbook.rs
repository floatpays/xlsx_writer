//! Building a rust_xlsxwriter `Workbook` from the decoded sheets.

use rust_xlsxwriter::{DocProperties, Workbook};
use rustler::Term;

use crate::error::{WriteError, WriteResult};
use crate::protocol::{Instruction, WorkbookProperties};
use crate::worksheet::{apply_instruction, finish_sheet, SheetState};

pub fn write_workbook(sheets: Vec<Term>, properties: Term) -> WriteResult<Vec<u8>> {
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
