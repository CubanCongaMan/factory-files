# Optical Televiewer Image Log Report Generator

Automated generation of Optical Televiewer Image Log reports from downhole camera images.

## Features

- Auto-crop downhole log images into 3-foot depth intervals
- Pull header data from Supabase PostgreSQL database
- Generate Word reports from templates
- Support for logs up to 66+ feet depth

## Installation

```bash
pip install -r requirements.txt
```

## Required Files

1. **optical_televiewer_report_generator.py** - Main Python script
2. **Template_18_Pages_No_Images_Rev_03.docx** - Word template (18 pages)
3. **Source image** - Full-depth downhole log JPG

## Supabase Database Setup

Create table in Supabase SQL Editor:

```sql
CREATE TABLE optical_televiewer_logs (
    id SERIAL PRIMARY KEY,
    vb_id_txt VARCHAR(15) NOT NULL,
    north_txt VARCHAR(15),
    easti_txt VARCHAR(15),
    stat_num VARCHAR(10),
    ground_elev_num NUMERIC(6,1),
    column_panel_txt VARCHAR(30),
    column_panel_joint_txt VARCHAR(30),
    ct_txt VARCHAR(10),
    drill_date DATE,
    drill_by_txt VARCHAR(15),
    op_tv_logger VARCHAR(25),
    op_tv_date DATE,
    image_file_path TEXT,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
```

## Usage

### Command Line

```bash
python optical_televiewer_report_generator.py \
    --vb-id H7-VB-04 \
    --image path/to/source_log.jpg \
    --template Template_18_Pages_No_Images_Rev_03.docx \
    --output-dir ./output \
    --api-key "your-supabase-service-role-key"
```

### As a Module

```python
from optical_televiewer_report_generator import generate_report

report_path = generate_report(
    vb_id="H7-VB-04",
    source_image_path="path/to/source_log.jpg",
    template_path="Template_18_Pages_No_Images_Rev_03.docx",
    output_dir="./output",
    service_role_key="your-supabase-service-role-key"
)
```

### Individual Functions

```python
from optical_televiewer_report_generator import (
    crop_downhole_log,
    query_log_data,
    insert_log_data
)

# Crop images only
image_files = crop_downhole_log("source.jpg", "./cropped")

# Query database
log_data = query_log_data("H7-VB-04", service_role_key)

# Insert new log data
new_log = {
    "vb_id_txt": "H7-VB-05",
    "north_txt": "779359.14",
    # ... other fields
}
insert_log_data(new_log, service_role_key)
```

## Configuration

Key parameters in the script (calibrated for standard logs):

| Parameter | Value | Description |
|-----------|-------|-------------|
| PIXELS_PER_FOOT | 360 | Scale calibration |
| HEADER_HEIGHT | 110 | Horizontal axis header (pixels) |
| RIGHT_CROP_X | 685 | Width crop boundary (pixels) |
| IMAGE_HEIGHT_INCHES | 7.0 | Image height in Word doc |

## Database Fields

| Field | Type | Description |
|-------|------|-------------|
| vb_id_txt | VARCHAR | Verification Boring ID |
| north_txt | VARCHAR | Northing coordinate |
| easti_txt | VARCHAR | Easting coordinate |
| stat_num | VARCHAR | Station number |
| ground_elev_num | NUMERIC | Ground elevation |
| column_panel_txt | VARCHAR | Column/Panel |
| column_panel_joint_txt | VARCHAR | Column/Panel joint |
| ct_txt | VARCHAR | CT Number |
| drill_date | DATE | Drill date |
| drill_by_txt | VARCHAR | Drilled by |
| op_tv_logger | VARCHAR | TV Logger operator |
| op_tv_date | DATE | TV logging date |

## Output

- **Cropped images**: `{output_dir}/cropped_images/{basename}_{depth}ft.jpg`
- **Word report**: `{output_dir}/Optical_Televiewer_Report_{vb_id}.docx`

## Supabase Connection

- URL: `https://oebbsdcdnnzqwdoacyyz.supabase.co`
- Requires service_role key for full access
