import aspose.words as aw

def create_report():
    # Initialize the document and builder
    doc = aw.Document()
    builder = aw.DocumentBuilder(doc)

    # Set global font and formatting
    builder.font.name = "Arial"
    builder.font.size = 12

    # Add the header
    builder.paragraph_format.alignment = aw.ParagraphAlignment.CENTER
    builder.writeln("FMG Solomon Firetail Stacker SK802")
    builder.writeln("Post Shut Report")

    builder.paragraph_format.alignment = aw.ParagraphAlignment.LEFT
    builder.writeln("Work Order: 061")
    builder.writeln("Document: BMS-03-REP-061_FT24OP11")

    builder.insert_break(aw.BreakType.PARAGRAPH_BREAK)

    # Add Document Revision Table
    builder.writeln("Document Revision")
    table = builder.start_table()

    # Add header row
    builder.insert_cell()
    builder.write("Rev")
    builder.insert_cell()
    builder.write("Date")
    builder.insert_cell()
    builder.write("Prepared")
    builder.insert_cell()
    builder.write("Approved")
    builder.insert_cell()
    builder.write("Comments")
    builder.end_row()

    # Add rows
    revisions = [
        ["0", "20/11/2024", "CN", "DR", "Issued for Internal review"],
        ["1", "28/11/2024", "CN", "DR", "Issued to Client"]
    ]

    for rev in revisions:
        for cell in rev:
            builder.insert_cell()
            builder.write(cell)
        builder.end_row()

    builder.end_table()
    builder.insert_break(aw.BreakType.PARAGRAPH_BREAK)

    # Add company info
    builder.writeln("Balance Machine Services")
    builder.writeln("ABN: 18 663 508 576")
    builder.writeln("https://bmservices.com.au/")
    builder.writeln("Unit 1, 64 Baile Rd")
    builder.writeln("Canning Vale WA 6155")
    builder.writeln("Contact: planning@bmservices.com.au")
    builder.insert_break(aw.BreakType.PAGE_BREAK)

    # Add Contents Section
    builder.writeln("CONTENTS")
    contents = [
        "1\tIntroduction",
        "2\tSafety",
        "3\tDisclaimers and Limitations",
        "4\tShutdown Scopes / Work Orders",
        "4.1\tSummary",
        "5\tUncompleted Work Scopes",
        "5.1\t5Y ME OFF Repl Luff HPU Pump Assy SK802",
        "5.2\t5Y ME OFF Repl Relief Vlv Luff Cyl SK802",
        "5.3\t5Y ME OFF Replace Luff Pressure Transducers",
        "6\tScopes Completed",
        "6.1\t52W SK802 Replace Hard Skirts & Soft Skirts (Break in Job)",
        "6.2\t52W ME OFF Adjust Cable Reel Torque Limit SK802",
        "7\tFurther Recommendations / Actions",
        "7.1\tParts Supply",
        "7.2\tPersonnel",
        "7.3\tShut Support",
        "7.4\tSafety",
    ]

    for item in contents:
        builder.writeln(item)

    builder.insert_break(aw.BreakType.PAGE_BREAK)

    # Add Section Placeholders with Headings
    sections = [
        "Introduction",
        "Safety",
        "Disclaimers and Limitations",
        "Shutdown Scopes / Work Orders",
        "Uncompleted Work Scopes",
        "Scopes Completed",
        "Further Recommendations / Actions"
    ]

    for section in sections:
        builder.writeln(section)
        builder.insert_break(aw.BreakType.PARAGRAPH_BREAK)
        builder.writeln("[Content Placeholder for {}]".format(section))
        builder.insert_break(aw.BreakType.PAGE_BREAK)

    # Add Sample Tables for Completed and Uncompleted Scopes
    builder.writeln("Scopes Completed")
    table = builder.start_table()

    # Table Headers
    headers = ["WORK ORDER#", "SCOPE", "Complete", "Incomplete", "Further works required", "Date Completed"]
    for header in headers:
        builder.insert_cell()
        builder.write(header)
    builder.end_row()

    # Table Rows Placeholder
    rows = [
        ["2200833463", "Replace Hard Skirts", "☒", "☐", "☐", "11/11/2024"],
        ["2200953011", "Maint HS Coupl Assy", "☒", "☐", "☐", "8/11/2024"]
    ]

    for row in rows:
        for cell in row:
            builder.insert_cell()
            builder.write(cell)
        builder.end_row()

    builder.end_table()

    # Add Recommendations Section
    builder.writeln("Further Recommendations / Actions")
    builder.writeln("Parts Supply")
    builder.writeln("[Details Placeholder]")
    builder.writeln("Personnel")
    builder.writeln("[Details Placeholder]")
    builder.writeln("Shut Support")
    builder.writeln("[Details Placeholder]")
    builder.writeln("Safety")
    builder.writeln("[Details Placeholder]")

    # Save the document
    output_path = "replicated_report_full.docx"
    doc.save(output_path)
    print(f"Document created and saved as {output_path}")

if __name__ == "__main__":
    create_report()
