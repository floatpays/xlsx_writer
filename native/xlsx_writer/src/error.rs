use std::fmt;

use rust_xlsxwriter::XlsxError;
use rustler::types::tuple::get_tuple;
use rustler::{Binary, ListIterator, Term, TermType};

/// Error returned to Elixir as the `reason` in `{:error, reason}`.
#[derive(Debug, PartialEq)]
pub struct WriteError(pub String);

impl WriteError {
    pub fn new(message: impl Into<String>) -> Self {
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

pub type WriteResult<T> = Result<T, WriteError>;

/// The longest description of a term that goes into an error message.
const MAX_DESCRIPTION_LENGTH: usize = 200;
/// Strings up to this many bytes are shown; longer binaries (e.g. images) are
/// only described by their size.
const MAX_STRING_LENGTH: usize = 64;
/// How many elements of a list are shown.
const MAX_LIST_ELEMENTS: usize = 3;

/// A short, bounded description of a term for an error message: small values
/// are shown in Erlang syntax, binaries only when they are short strings, and
/// the result is cut off after `MAX_DESCRIPTION_LENGTH` characters. Walking
/// the term stops there too, so a large sheet is not traversed.
pub fn describe_term(term: Term) -> String {
    let mut description = String::new();
    describe_into(term, &mut description);

    if description.len() > MAX_DESCRIPTION_LENGTH {
        let mut end = MAX_DESCRIPTION_LENGTH;
        while !description.is_char_boundary(end) {
            end -= 1;
        }
        description.truncate(end);
        description.push_str("...");
    }

    description
}

fn describe_into(term: Term, description: &mut String) {
    if description.len() > MAX_DESCRIPTION_LENGTH {
        return;
    }

    match term.get_type() {
        TermType::Atom => match term.atom_to_string() {
            Ok(atom) => description.push_str(&atom),
            Err(_) => description.push_str("atom"),
        },
        TermType::Integer => match term.decode::<i64>() {
            Ok(integer) => description.push_str(&integer.to_string()),
            Err(_) => description.push_str("integer"),
        },
        TermType::Float => match term.decode::<f64>() {
            Ok(float) => description.push_str(&float.to_string()),
            Err(_) => description.push_str("float"),
        },
        TermType::Binary => match term.decode::<Binary>() {
            Ok(binary) => describe_binary(binary.as_slice(), description),
            Err(_) => description.push_str("binary"),
        },
        TermType::Tuple => match get_tuple(term) {
            Ok(elements) => {
                description.push('{');
                for (index, element) in elements.into_iter().enumerate() {
                    if index > 0 {
                        description.push(',');
                    }
                    describe_into(element, description);
                    if description.len() > MAX_DESCRIPTION_LENGTH {
                        return;
                    }
                }
                description.push('}');
            }
            Err(_) => description.push_str("tuple"),
        },
        TermType::List => match term.decode::<ListIterator>() {
            Ok(elements) => {
                description.push('[');
                for (index, element) in elements.enumerate() {
                    if index == MAX_LIST_ELEMENTS {
                        description.push_str(",...");
                        break;
                    }
                    if index > 0 {
                        description.push(',');
                    }
                    describe_into(element, description);
                    if description.len() > MAX_DESCRIPTION_LENGTH {
                        return;
                    }
                }
                description.push(']');
            }
            Err(_) => description.push_str("list"),
        },
        TermType::Map => description.push_str("#{...}"),
        TermType::Fun => description.push_str("fun"),
        TermType::Pid => description.push_str("pid"),
        TermType::Port => description.push_str("port"),
        TermType::Ref => description.push_str("reference"),
        TermType::Unknown => description.push_str("term"),
    }
}

fn describe_binary(bytes: &[u8], description: &mut String) {
    match std::str::from_utf8(bytes) {
        Ok(text) if bytes.len() <= MAX_STRING_LENGTH => {
            description.push_str(&format!("{text:?}"));
        }
        _ => description.push_str(&format!("<<{} bytes>>", bytes.len())),
    }
}
