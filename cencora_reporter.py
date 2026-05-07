# Databricks notebook source
# MAGIC %md
# MAGIC ## Data Export Function - Usage Guide
# MAGIC
# MAGIC ### Overview
# MAGIC This notebook exports Unity Catalog table or SQL query results to delimited files in UC Volumes.
# MAGIC
# MAGIC ### Parameters (Widgets)
# MAGIC
# MAGIC | Parameter | Type | Description | Default |
# MAGIC |-----------|------|-------------|----------|
# MAGIC | `source` | Text | Table name (catalog.schema.table) or SQL query | Required |
# MAGIC | `delimiter_type` | Dropdown | Delimiter type: comma, pipe, or tilde | comma |
# MAGIC | `output_path` | Text | UC Volume path (e.g., /Volumes/catalog/schema/volume/) | Required |
# MAGIC | `include_header` | Dropdown | Include column headers (true/false) | true |
# MAGIC | `blank_rows_top` | Text | Number of blank rows at top | 0 |
# MAGIC | `include_footer` | Dropdown | Include footer with record count (true/false) | true |
# MAGIC | `output_filename` | Text | Output file name | output.txt |
# MAGIC
# MAGIC ### Example Usage
# MAGIC
# MAGIC #### Example 1: Export a table with comma delimiter
# MAGIC ```
# MAGIC source: main.default.my_table
# MAGIC delimiter_type: comma
# MAGIC output_path: /Volumes/main/default/my_volume/
# MAGIC include_header: true
# MAGIC blank_rows_top: 0
# MAGIC include_footer: true
# MAGIC output_filename: export.csv
# MAGIC ```
# MAGIC
# MAGIC #### Example 2: Export SQL query with pipe delimiter
# MAGIC ```
# MAGIC source: SELECT customer_id, order_date, total FROM main.default.orders WHERE order_date >= '2024-01-01'
# MAGIC delimiter_type: pipe
# MAGIC output_path: /Volumes/main/default/exports/
# MAGIC include_header: true
# MAGIC blank_rows_top: 2
# MAGIC include_footer: true
# MAGIC output_filename: orders_export.txt
# MAGIC ```
# MAGIC
# MAGIC #### Example 3: Export with tilde delimiter, no header
# MAGIC ```
# MAGIC source: main.default.products
# MAGIC delimiter_type: tilde
# MAGIC output_path: /Volumes/main/default/data_exports/
# MAGIC include_header: false
# MAGIC blank_rows_top: 0
# MAGIC include_footer: false
# MAGIC output_filename: products.txt
# MAGIC ```
# MAGIC
# MAGIC ### Running as a Databricks Job
# MAGIC
# MAGIC 1. Create a new job in Databricks
# MAGIC 2. Add this notebook as a task
# MAGIC 3. Configure parameters in the job definition
# MAGIC 4. Schedule or trigger the job as needed
# MAGIC
# MAGIC ### Notes
# MAGIC
# MAGIC * The function collects data to the driver, so be mindful of dataset size
# MAGIC * For very large datasets, consider using Spark's native write methods with coalesce(1)
# MAGIC * Ensure the output UC Volume exists and you have write permissions
# MAGIC * The function overwrites existing files with the same name

# COMMAND ----------

# DBTITLE 1,Setup Parameters with Widgets
# Create widgets for job parameters
dbutils.widgets.text("source", "", "Source (table name or SQL query)")
dbutils.widgets.dropdown("delimiter_type", "comma", ["comma", "pipe", "tilde"], "Delimiter Type")
dbutils.widgets.text("output_path", "", "Output Path (UC Volume path)")
dbutils.widgets.dropdown("include_header", "true", ["true", "false"], "Include Header")
dbutils.widgets.text("blank_rows_top", "0", "Blank Rows at Top")
dbutils.widgets.dropdown("include_footer", "true", ["true", "false"], "Include Footer with Count")
dbutils.widgets.text("output_filename", "output.txt", "Output Filename")

# COMMAND ----------

# DBTITLE 1,Define Export Function
def export_to_delimited_file(
    source: str,
    output_path: str,
    delimiter_type: str = "comma",
    include_header: bool = True,
    blank_rows_top: int = 0,
    include_footer: bool = True,
    output_filename: str = "output.txt"
):
    """
    Export Unity Catalog table or SQL query results to a delimited file.
    
    Parameters:
    -----------
    source : str
        Either a fully qualified table name (catalog.schema.table) or a SQL query
    output_path : str
        UC Volume path where the file will be saved (e.g., /Volumes/catalog/schema/volume/)
    delimiter_type : str
        Type of delimiter: 'comma', 'pipe', or 'tilde'
    include_header : bool
        Whether to include column headers in the output
    blank_rows_top : int
        Number of blank rows to add at the top of the file
    include_footer : bool
        Whether to include a footer row with record count
    output_filename : str
        Name of the output file
    
    Returns:
    --------
    dict : Summary information about the export
    """
    
    # Map delimiter types to actual delimiters
    delimiter_map = {
        "comma": ",",
        "pipe": "|",
        "tilde": "~"
    }
    
    delimiter = delimiter_map.get(delimiter_type.lower(), ",")
    
    # Determine if source is a table or SQL query
    if (source.strip().upper().startswith("SELECT") or source.strip().upper().startswith("WITH")):
        # It's a SQL query
        df = spark.sql(source)
        source_type = "SQL Query"
    else: # additional error checking required
        # Assume it's a table name
        df = spark.table(source)
        source_type = "Table"
    
    # Collect data to driver (be cautious with large datasets)
    data = df.collect()
    columns = df.columns
    
    # Construct full output path
    full_output_path = f"{output_path.rstrip('/')}/{output_filename}"
    
    # Build file content
    lines = []
    
    # Add blank rows at top
    for _ in range(blank_rows_top):
        lines.append("")
    
    # Add header row
    if include_header:
        header_line = delimiter.join(columns)
        lines.append(header_line)
    
    # Add data rows
    for row in data:
        # Convert row to list of strings, handling None values
        row_values = [str(val) if val is not None else "" for val in row]
        data_line = delimiter.join(row_values)
        lines.append(data_line)
    
    # Add footer with record count
    if include_footer:
        # Get record count
        record_count = df.count()
        footer_line = f"Total Records: {record_count}"
        lines.append(footer_line)
    
    # Join all lines with newline
    file_content = "\n".join(lines)
    
    # Write to file using dbutils
    dbutils.fs.put(full_output_path, file_content, overwrite=True)
    
    # Return summary
    summary = {
        "source_type": source_type,
        "source": source,
        "output_path": full_output_path,
        "delimiter": delimiter,
        "record_count": record_count,
        "header_included": include_header,
        "blank_rows_top": blank_rows_top,
        "footer_included": include_footer,
        "total_lines": len(lines)
    }
    
    return summary

# COMMAND ----------

# DBTITLE 1,Execute Export with Widget Parameters
# Get widget values
source = dbutils.widgets.get("source")
delimiter_type = dbutils.widgets.get("delimiter_type")
output_path = dbutils.widgets.get("output_path")
include_header = dbutils.widgets.get("include_header").lower() == "true"
blank_rows_top = int(dbutils.widgets.get("blank_rows_top"))
include_footer = dbutils.widgets.get("include_footer").lower() == "true"
output_filename = dbutils.widgets.get("output_filename")

# Validate required parameters
if not source:
    raise ValueError("Source parameter is required (table name or SQL query)")

if not output_path:
    raise ValueError("Output path parameter is required (UC Volume path)")

# Execute the export
print(f"Starting export...")
print(f"Source: {source}")
print(f"Delimiter: {delimiter_type}")
print(f"Output: {output_path}/{output_filename}")
print("-" * 80)

summary = export_to_delimited_file(
    source=source,
    output_path=output_path,
    delimiter_type=delimiter_type,
    include_header=include_header,
    blank_rows_top=blank_rows_top,
    include_footer=include_footer,
    output_filename=output_filename
)

# Display summary
print("\n✓ Export completed successfully!")
print("\nSummary:")
for key, value in summary.items():
    print(f"  {key}: {value}")

# COMMAND ----------

# MAGIC %md
# MAGIC ```
# MAGIC WITH heart_failure_patients AS (
# MAGIC   SELECT
# MAGIC     co.PERSON_ID,
# MAGIC     MIN(co.CONDITION_START_DATE) AS condition_start_date
# MAGIC   FROM
# MAGIC     `tenbosch`.`omop`.`condition_occurrence` co
# MAGIC   WHERE
# MAGIC     co.CONDITION_CONCEPT_ID IN (
# MAGIC       SELECT
# MAGIC         ca.descendant_concept_id
# MAGIC       FROM
# MAGIC         `tenbosch`.`omop`.`concept_ancestor` ca
# MAGIC       WHERE
# MAGIC         ca.ancestor_concept_id = 4229440
# MAGIC     )
# MAGIC   GROUP BY
# MAGIC     co.PERSON_ID
# MAGIC ),
# MAGIC alteplase_concept AS (
# MAGIC   SELECT
# MAGIC     concept_id
# MAGIC   FROM
# MAGIC     `tenbosch`.`omop`.`concept`
# MAGIC   WHERE
# MAGIC     concept_name ILIKE '%alteplase 100 MG Injection%'
# MAGIC ),
# MAGIC alteplase_descendants AS (
# MAGIC   SELECT
# MAGIC     ca.descendant_concept_id
# MAGIC   FROM
# MAGIC     `tenbosch`.`omop`.`concept_ancestor` ca
# MAGIC       JOIN alteplase_concept ac
# MAGIC         ON ca.ancestor_concept_id = ac.concept_id
# MAGIC ),
# MAGIC alteplase_patients AS (
# MAGIC   SELECT DISTINCT
# MAGIC     de.PERSON_ID
# MAGIC   FROM
# MAGIC     `tenbosch`.`omop`.`drug_exposure` de
# MAGIC   WHERE
# MAGIC     de.DRUG_CONCEPT_ID IN (
# MAGIC       SELECT
# MAGIC         descendant_concept_id
# MAGIC       FROM
# MAGIC         alteplase_descendants
# MAGIC     )
# MAGIC )
# MAGIC SELECT
# MAGIC   hfp.PERSON_ID,
# MAGIC   hfp.condition_start_date
# MAGIC FROM
# MAGIC   heart_failure_patients hfp
# MAGIC     LEFT JOIN alteplase_patients ap
# MAGIC       ON hfp.PERSON_ID = ap.PERSON_ID
# MAGIC WHERE
# MAGIC   ap.PERSON_ID IS NULL
# MAGIC
# MAGIC ```