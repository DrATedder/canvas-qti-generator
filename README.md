![Banner Image](https://github.com/DrATedder/canvas-qti-generator/blob/main/CANVAS_QTI_generator_header.png)
# Canvas QTI Generator

A Java desktop application that converts a structured Excel (.xls) file into a Canvas-compatible IMS QTI 1.2 question bank ZIP package.

## Features

- Converts .xls multiple-choice spreadsheets into QTI 1.2
- Automatically generates imsmanifest.xml
- Names output ZIP after assessment title
- Handles numeric and string Excel cell types safely
- Skips blank rows
- Generates unique IDs if missing
- Executable fat JAR build

## Excel Format

For an example input file, see [Example_sheet.xls](Example_sheet.xls).

| Cell | Purpose |
|------|----------|
| B1 | Assessment Identity |
| B2 | Assessment Title |
| Row 4+ | Questions |

Columns:
- A = Question Text
- B = Item Identity
- C = Item Title
- D = Option A
- E = Option B
- F = Option C
- G = Option D
- H = Correct Answer (A/B/C/D)

## Build

**Note.** A pre-compiled, executable version of the app is available in the 'Releases' tab.

Requires:
- Java 17+
- Maven 3.9+

## Compile

```bash
mvn clean package
```

## Run

```bash
java -jar target/canvas-qti-generator-1.0.jar
```


## Import into Canvas

1. Course → Settings
2. Import Course Content
3. Content Type → QTI .zip file
4. Upload generated ZIP

## License

MIT License

![Java](https://img.shields.io/badge/Java-17-blue)
![Maven](https://img.shields.io/badge/Maven-3.9+-red)
![License](https://img.shields.io/badge/License-MIT-green)
