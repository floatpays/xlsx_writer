use std::fmt;

use rust_xlsxwriter::XlsxError;

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
