use std::borrow::Cow;
use std::{path::Path, fs::File, borrow::BorrowMut};
use std::io::Write;

use calamine::vba::VbaProject;
use calamine::{Xls, open_workbook, Reader, DataType, Range, Xlsx};
use clap::Parser;

#[derive(Parser)]
struct Args {
    /// Xls File to Open
    #[clap(parse(from_os_str), short='i', long)]
    xls_fname: std::path::PathBuf,

    /// Export Folder to Storage Source
    #[clap(parse(from_os_str), short='o', long)]
    export_folder: std::path::PathBuf,

    /// Delimiter for sheets cells
    #[clap(short='d', long, default_value="\t")]
    sheet_delimiter: String,

    /// Extension for sheet file (txt/csv)
    #[clap(short='x', long, default_value="txt")]
    sheet_extension: String,
}

fn export_sheets
(
    export_folder: &std::path::PathBuf,
    sheet_delimiter: &String,
    sheet_extension: &String,
    worksheets: &Vec<(String, Range<DataType>)>
) {
    
    // go through all worksheets and export as txt
    for (name, data) in worksheets {
        let mut path = export_folder.to_owned();
        path.push(name.to_owned() + "." + sheet_extension.to_owned().as_str());
        let mut file = File::create(path).unwrap();

        if data.start().is_none() || data.end().is_none() {
            continue;
        }

        let mut row = 0;
        let mut col = 0;

        // go through all rows
        while row < data.end().unwrap().0 {
            // reset column position
            col = 0;
            while col < data.end().unwrap().1 {
                // get out value as a String
                let val = if let Some(val) = data.get_value((row, col)) {
                    format!("{}", val)
                } else {
                    "".to_string()
                };

                // Write the value to the file
                write!(file, "{}{}", val, sheet_delimiter).unwrap();

                // increment column counter
                col += 1;
            }

            // write out a newline
            writeln!(file, "").unwrap();

            // increment the row counter
            row += 1;
        }
    }
}

fn export_vba_proj
(
    export_folder: &std::path::PathBuf,
    sheet_names: &Vec<String>,
    proj: Cow<VbaProject>
) {
    // go through all vba modules/classes/forms and export as bas/cls/frm
    let modules = proj.get_module_names().to_owned();
    for module in modules {
        let name: String = module.to_string().to_owned();
        if !sheet_names.contains(&name) {

            if let Ok(contents) = proj.get_module(&name) {
                let mut path = export_folder.to_owned();
                if contents.contains("\nAttribute VB_Base") {
                    // class
                    path.push(name.to_owned() + ".cls");
                } else {
                    // export as bas file
                    path.push(name.to_owned() + ".bas");
                }
                let mut file = File::create(path).unwrap();

                if let Ok(contents) = proj.get_module(&name) {
                    write!(file, "{}", contents).unwrap();
                }
            }
        }
    }
}

fn main() {
    let args = Args::parse();

    println!("Executing with following arguments:");
    println!("xls_fname:     {}", args.xls_fname.to_str().unwrap());
    println!("export_folder: {}", args.export_folder.to_str().unwrap());

    let mut sheet_names = Vec::new();
    let mut worksheets = Vec::new();
    let mut proj = None;

    // create export folder if it doesn't exist
    if !Path::exists(&args.export_folder.to_owned()) {
        std::fs::create_dir_all(args.export_folder.to_owned()).unwrap()
    }

    if args.xls_fname.ends_with("xls") {
        // xls file
        let mut workbook: Xls<_> = open_workbook(args.xls_fname).unwrap();
        sheet_names = workbook.sheet_names().to_vec();
        worksheets = workbook.worksheets();
        proj = if let Some(Ok(proj)) = workbook.vba_project() {
            Some(proj)
        } else {
            None
        };

        // grab the sheet names, we'll need it
        sheet_names.push("ThisWorkbook".to_string());
    
        export_sheets(&args.export_folder, &args.sheet_delimiter, &args.sheet_extension, &worksheets);
        if let Some(proj) = proj {export_vba_proj(&args.export_folder, &sheet_names, proj);}
    } else {
        // xlsm file
        let mut workbook: Xlsx<_> = open_workbook(args.xls_fname).unwrap();
        sheet_names = workbook.sheet_names().to_vec();
        worksheets = workbook.worksheets();
        proj = if let Some(Ok(proj)) = workbook.vba_project() {
            Some(proj)
        } else {
            None
        };

        // grab the sheet names, we'll need it
        sheet_names.push("ThisWorkbook".to_string());
    
        export_sheets(&args.export_folder, &args.sheet_delimiter, &args.sheet_extension, &worksheets);
        if let Some(proj) = proj {export_vba_proj(&args.export_folder, &sheet_names, proj);}
    }

}
