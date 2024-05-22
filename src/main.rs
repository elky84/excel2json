use calamine::{open_workbook, Reader, Xlsx, DataType};
use serde_json::json;
use std::fs::{File, create_dir_all};
use std::io::Write;
use chrono::{DateTime, Utc, TimeZone};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // 지정된 폴더가 없는 경우 생성합니다.
    let output_folder = "output";
    create_dir_all(output_folder)?;

    // 엑셀 파일을 엽니다.
    let mut workbook: Xlsx<_> = open_workbook("data.xlsx")?;

    // 시트 이름을 미리 수집합니다.
    let sheet_names: Vec<String> = workbook.sheet_names().iter().map(|s| s.to_string()).collect();

    // 모든 시트를 순회합니다.
    for sheet_name in sheet_names {
        if let Some(Ok(range)) = workbook.worksheet_range(&sheet_name) {
            // 데이터를 저장할 벡터를 정의합니다.
            let mut column_names = Vec::new();
            let mut data_types = Vec::new();
            let mut categories = Vec::new();
            let mut data = Vec::new();

            for (i, row) in range.rows().enumerate() {
                match i {
                    0 => {
                        // 첫 번째 행은 컬럼명
                        column_names = row.iter().map(|cell| cell.to_string()).collect();
                    }
                    1 => {
                        // 두 번째 행은 데이터 타입
                        data_types = row.iter().map(|cell| cell.to_string()).collect();
                    }
                    2 => {
                        // 세 번째 행은 category
                        categories = row.iter().map(|cell| cell.to_string()).collect();
                    }
                    _ => {
                        // 나머지 행은 실제 데이터
                        let row_data: Vec<String> = row.iter().map(|cell| cell_to_string(cell)).collect();
                        data.push(row_data);
                    }
                }
            }

            // zip3 함수에 전달할 벡터를 생성합니다.
            let column_names_vec = column_names.iter().map(|s| s.to_string()).collect::<Vec<_>>();
            let data_types_vec = data_types.iter().map(|s| s.to_string()).collect::<Vec<_>>();
            let categories_vec = categories.iter().map(|s| s.to_string()).collect::<Vec<_>>();

            // C# 클래스 생성
            let mut class_code = String::new();
            class_code.push_str(&format!("public class {}\n{{\n", sheet_name));

            for (name, data_type, category) in zip3(column_names_vec.clone(), data_types_vec.clone(), categories_vec.clone()) {
                let cs_type = match data_type.as_str() {
                    "int" => "int",
                    "string" => "string",
                    _ => "string", // 기본적으로 string으로 처리합니다.
                };
                class_code.push_str(&format!("    public {} {} {{ get; set; }} // Category: {}\n", cs_type, name, category));
            }

            class_code.push_str("}\n");

            // 생성된 C# 코드를 파일에 작성합니다.
            let file_name = format!("{}/{}.cs", output_folder, sheet_name);
            let mut file = File::create(&file_name)?;
            file.write_all(class_code.as_bytes())?;

            println!("C# 코드가 생성되었습니다: {}", file_name);

            // JSON 데이터 생성
            let mut json_data = Vec::new();
            for row in data {
                let mut json_row = serde_json::Map::new();
                for (name, value) in column_names_vec.iter().zip(row.iter()) {
                    json_row.insert(name.clone(), json!(value));
                }
                json_data.push(json_row);
            }

            // JSON 파일 작성
            let json_file_name = format!("{}/{}.json", output_folder, sheet_name);
            let json_file = File::create(&json_file_name)?;
            serde_json::to_writer_pretty(json_file, &json_data)?;

            println!("JSON 데이터가 생성되었습니다: {}", json_file_name);
        }
    }

    Ok(())
}

// zip3 함수는 세 개의 벡터를 함께 이터레이트하기 위한 도우미 함수입니다.
fn zip3<A, B, C>(a: Vec<A>, b: Vec<B>, c: Vec<C>) -> impl Iterator<Item = (A, B, C)> {
    a.into_iter().zip(b).zip(c).map(|((a, b), c)| (a, b, c))
}

fn cell_to_string(cell: &DataType) -> String {
    match cell {
        DataType::Empty => "".to_string(),
        DataType::String(s) => s.clone(),
        DataType::Float(f) => f.to_string(),
        DataType::Int(i) => i.to_string(),
        DataType::Bool(b) => b.to_string(),
        DataType::Error(e) => format!("Error: {:?}", e),
        DataType::DateTime(dt) => {
            DateTime::from_timestamp(*dt as i64, 0)
                .expect("Failed to convert to DateTime")
                .to_string() // 값을 추출하여 문자열로 변환
        }
    }
}
