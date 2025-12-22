use roxmltree::{Document, Node};
use std::ops::RangeBounds;

const NAMESPACE_SS: &str = "urn:schemas-microsoft-com:office:spreadsheet";

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum ParseError {
    XmlError(roxmltree::Error),
    ExcelFormatError(String),
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum DataType {
    String,
    Number,
    DateTime,
    Boolean,
    Error,
    Unknown,
}

impl From<&str> for DataType {
    fn from(s: &str) -> Self {
        match s {
            "String" => DataType::String,
            "Number" => DataType::Number,
            "DateTime" => DataType::DateTime,
            "Boolean" => DataType::Boolean,
            "Error" => DataType::Error,
            _ => DataType::Unknown,
        }
    }
}

impl From<Option<&str>> for DataType {
    fn from(opt: Option<&str>) -> Self {
        match opt {
            Some(s) => DataType::from(s),
            None => DataType::Unknown,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Workbook {
    pub sheets: Vec<Sheet>,
}

impl TryFrom<&str> for Workbook {
    type Error = ParseError;

    fn try_from(s: &str) -> Result<Self, Self::Error> {
        let doc = Document::parse(s).map_err(ParseError::XmlError)?;
        let root = doc
            .descendants()
            .find(|n| n.has_tag_name("Workbook"))
            .ok_or_else(|| ParseError::ExcelFormatError("Missing Workbook element".to_string()))?;

        let sheet_nodes = root
            .children()
            .filter(|n| {
                n.has_tag_name("Worksheet") && n.children().any(|c| c.has_tag_name("Table"))
            })
            .collect::<Vec<_>>();
        if sheet_nodes.is_empty() {
            return Err(ParseError::ExcelFormatError(
                "No worksheets with tables found".to_string(),
            ));
        }

        let sheets = sheet_nodes
            .into_iter()
            .flat_map(Sheet::try_from)
            .collect::<Vec<_>>();

        Ok(Workbook { sheets })
    }
}

impl Workbook {
    pub fn get_sheet_by_name(&self, name: &str) -> Option<&Sheet> {
        self.sheets.iter().find(|s| s.name == name)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Sheet {
    pub name: String,
    pub table: Vec<Row>,
}

impl TryFrom<Node<'_, '_>> for Sheet {
    type Error = ParseError;

    fn try_from(node: Node<'_, '_>) -> Result<Self, Self::Error> {
        let mut r_idx = 0;

        let name = node
            .attribute((NAMESPACE_SS, "Name"))
            .ok_or_else(|| {
                ParseError::ExcelFormatError("Worksheet missing Name attribute".to_string())
            })?
            .to_string();

        let table = node
            .children()
            .find(|n| n.has_tag_name("Table"))
            .ok_or_else(|| ParseError::ExcelFormatError("Missing Table element".to_string()))?
            .children()
            .filter(|n| n.has_tag_name("Row"))
            .filter_map(|r| {
                let row = Row::try_from((r, r_idx)).ok()?;
                r_idx = row.index;
                Some(row)
            })
            .collect::<Vec<_>>();

        match table[..] {
            [] => Err(ParseError::ExcelFormatError(
                "Sheet contains no data".to_string(),
            )),
            _ => Ok(Sheet { name, table }),
        }
    }
}

impl Sheet {
    pub fn get_row_by_index(&self, index: usize) -> Option<&Row> {
        self.table.iter().find(|r| r.index == index)
    }

    pub fn get_rows_by_range(&self, range: impl RangeBounds<usize>) -> Vec<&Row> {
        self.table
            .iter()
            .filter(|r| range.contains(&r.index))
            .collect()
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Row {
    pub index: usize,
    pub cells: Vec<Cell>,
}

impl TryFrom<(Node<'_, '_>, usize)> for Row {
    type Error = ParseError;

    fn try_from((node, prev_idx): (Node<'_, '_>, usize)) -> Result<Self, Self::Error> {
        let index = match node.attribute((NAMESPACE_SS, "Index")) {
            Some(s) => s.parse().unwrap_or(prev_idx + 1),
            None => prev_idx + 1,
        };

        let cells = {
            let mut c_idx = 0;
            node.children()
                .filter(|n| n.has_tag_name("Cell"))
                .filter_map(|c| {
                    let cell = Cell::try_from((c, c_idx)).ok()?;
                    c_idx = cell.index;
                    Some(cell)
                })
                .collect::<Vec<_>>()
        };

        match cells[..] {
            [] => Err(ParseError::ExcelFormatError(
                "Row contains no cells".to_string(),
            )),
            _ => Ok(Row { index, cells }),
        }
    }
}

impl Row {
    pub fn get_cell_by_index(&self, index: usize) -> Option<&Cell> {
        self.cells.iter().find(|c| c.index == index)
    }

    pub fn get_cells_by_range(&self, range: impl RangeBounds<usize>) -> Vec<&Cell> {
        self.cells
            .iter()
            .filter(|c| range.contains(&c.index))
            .collect()
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Cell {
    pub index: usize,
    pub data_type: DataType,
    pub value: String,
}

impl TryFrom<(Node<'_, '_>, usize)> for Cell {
    type Error = ParseError;

    fn try_from((node, prev_idx): (Node<'_, '_>, usize)) -> Result<Self, Self::Error> {
        let data_node = node
            .children()
            .find(|n| n.has_tag_name("Data"))
            .ok_or_else(|| ParseError::ExcelFormatError("Cell missing Data element".to_string()))?;

        let value = data_node
            .text()
            .ok_or_else(|| ParseError::ExcelFormatError("Data element missing text".to_string()))?
            .trim()
            .to_string();

        let data_type = DataType::from(data_node.attribute((NAMESPACE_SS, "Type")));

        let index = match node.attribute((NAMESPACE_SS, "Index")) {
            Some(s) => s.parse().unwrap_or(prev_idx + 1),
            None => prev_idx + 1,
        };

        Ok(Cell {
            index,
            data_type,
            value,
        })
    }
}
