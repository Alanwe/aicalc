# AiCalc Function Reference

**Version:** 1.0  
**Last Updated:** October 19, 2025

---

## Overview

AiCalc includes 30+ built-in functions across multiple categories: Math, Text, DateTime, File, Directory, Table, Image, PDF, and AI. This document provides comprehensive reference for each function with syntax, parameters, examples, and return types.

---

## Table of Contents

- [Math Functions](#math-functions)
- [Text Functions](#text-functions)
- [DateTime Functions](#datetime-functions)
- [File Functions](#file-functions)
- [Directory Functions](#directory-functions)
- [Table Functions](#table-functions)
- [Image Functions](#image-functions)
- [PDF Functions](#pdf-functions)
- [AI Functions](#ai-functions)

---

## Math Functions

### SUM
**Description:** Adds a series of numbers.

**Syntax:** `=SUM(values...)`

**Parameters:**
- `values` (Number, variadic): Range of values to add

**Returns:** Number

**Examples:**
```
=SUM(A1:A10)          // Sum cells A1 through A10
=SUM(5, 10, 15)       // Returns 30
=SUM(B2, C2, D2)      // Sum three cells
```

---

### AVERAGE
**Description:** Calculates the average (mean) of a series of numbers.

**Syntax:** `=AVERAGE(values...)`

**Parameters:**
- `values` (Number, variadic): Range of values to average

**Returns:** Number

**Examples:**
```
=AVERAGE(A1:A5)       // Average of 5 cells
=AVERAGE(10, 20, 30)  // Returns 20
```

---

### COUNT
**Description:** Counts the number of numeric values in a range.

**Syntax:** `=COUNT(values...)`

**Parameters:**
- `values` (Number, variadic): Range of values to count

**Returns:** Number

**Examples:**
```
=COUNT(A1:A10)        // Count numeric cells in range
=COUNT(5, "text", 10) // Returns 2 (ignores text)
```

---

### MIN
**Description:** Returns the minimum value from a series of numbers.

**Syntax:** `=MIN(values...)`

**Parameters:**
- `values` (Number, variadic): Range of values

**Returns:** Number

**Examples:**
```
=MIN(A1:A10)          // Find smallest value
=MIN(5, 3, 9, 1)      // Returns 1
```

---

### MAX
**Description:** Returns the maximum value from a series of numbers.

**Syntax:** `=MAX(values...)`

**Parameters:**
- `values` (Number, variadic): Range of values

**Returns:** Number

**Examples:**
```
=MAX(A1:A10)          // Find largest value
=MAX(5, 3, 9, 1)      // Returns 9
```

---

### ROUND
**Description:** Rounds a number to a specified number of decimal places.

**Syntax:** `=ROUND(value, [digits])`

**Parameters:**
- `value` (Number): Number to round
- `digits` (Number, optional): Number of decimal places (default: 0)

**Returns:** Number

**Examples:**
```
=ROUND(3.14159, 2)    // Returns 3.14
=ROUND(42.7)          // Returns 43
=ROUND(A1, 3)         // Round A1 to 3 decimals
```

---

### ABS
**Description:** Returns the absolute value of a number.

**Syntax:** `=ABS(value)`

**Parameters:**
- `value` (Number): Number

**Returns:** Number

**Examples:**
```
=ABS(-5)              // Returns 5
=ABS(3.14)            // Returns 3.14
=ABS(A1)              // Absolute value of A1
```

---

### SQRT
**Description:** Returns the square root of a number.

**Syntax:** `=SQRT(value)`

**Parameters:**
- `value` (Number): Number

**Returns:** Number

**Examples:**
```
=SQRT(16)             // Returns 4
=SQRT(2)              // Returns 1.414...
=SQRT(A1)             // Square root of A1
```

---

### POWER
**Description:** Returns a number raised to a power.

**Syntax:** `=POWER(base, exponent)`

**Parameters:**
- `base` (Number): Base number
- `exponent` (Number): Exponent

**Returns:** Number

**Examples:**
```
=POWER(2, 3)          // Returns 8 (2³)
=POWER(10, 2)         // Returns 100
=POWER(A1, 2)         // Square A1
```

---

## Text Functions

### CONCAT
**Description:** Concatenates (joins) string values.

**Syntax:** `=CONCAT(values...)`

**Parameters:**
- `values` (Text, variadic): Values to concatenate

**Returns:** Text

**Examples:**
```
=CONCAT("Hello", " ", "World")  // Returns "Hello World"
=CONCAT(A1, ", ", A2)           // Join with comma
=CONCAT(A1:A3)                  // Join range of cells
```

---

### UPPER
**Description:** Converts text to uppercase.

**Syntax:** `=UPPER(text)`

**Parameters:**
- `text` (Text): Text to convert

**Returns:** Text

**Examples:**
```
=UPPER("hello")       // Returns "HELLO"
=UPPER(A1)            // Uppercase of A1
```

---

### LOWER
**Description:** Converts text to lowercase.

**Syntax:** `=LOWER(text)`

**Parameters:**
- `text` (Text): Text to convert

**Returns:** Text

**Examples:**
```
=LOWER("HELLO")       // Returns "hello"
=LOWER(A1)            // Lowercase of A1
```

---

### TRIM
**Description:** Removes leading and trailing whitespace from text.

**Syntax:** `=TRIM(text)`

**Parameters:**
- `text` (Text): Text to trim

**Returns:** Text

**Examples:**
```
=TRIM("  hello  ")    // Returns "hello"
=TRIM(A1)             // Trim whitespace from A1
```

---

### LEN
**Description:** Returns the length (character count) of a text string.

**Syntax:** `=LEN(text)`

**Parameters:**
- `text` (Text): Text to measure

**Returns:** Number

**Examples:**
```
=LEN("Hello")         // Returns 5
=LEN(A1)              // Length of A1
```

---

### REPLACE
**Description:** Replaces occurrences of text within a string.

**Syntax:** `=REPLACE(text, old_text, new_text)`

**Parameters:**
- `text` (Text): Original text
- `old_text` (Text): Text to replace
- `new_text` (Text): Replacement text

**Returns:** Text

**Examples:**
```
=REPLACE("Hello World", "World", "Earth")  // Returns "Hello Earth"
=REPLACE(A1, "old", "new")                 // Replace in A1
```

---

### SPLIT
**Description:** Splits text into parts using a delimiter.

**Syntax:** `=SPLIT(text, [delimiter])`

**Parameters:**
- `text` (Text): Text to split
- `delimiter` (Text, optional): Separator character (default: ",")

**Returns:** Text (comma-separated parts)

**Examples:**
```
=SPLIT("a,b,c")              // Returns "a, b, c"
=SPLIT("a|b|c", "|")         // Returns "a, b, c"
=SPLIT(A1, ";")              // Split A1 by semicolon
```

---

## DateTime Functions

### NOW
**Description:** Returns the current date and time.

**Syntax:** `=NOW()`

**Parameters:** None

**Returns:** DateTime

**Examples:**
```
=NOW()                // Returns current date and time
```

---

### TODAY
**Description:** Returns the current date (without time).

**Syntax:** `=TODAY()`

**Parameters:** None

**Returns:** DateTime

**Examples:**
```
=TODAY()              // Returns current date
```

---

### DATE
**Description:** Creates a date from year, month, and day components.

**Syntax:** `=DATE(year, month, day)`

**Parameters:**
- `year` (Number): Year (e.g., 2025)
- `month` (Number): Month (1-12)
- `day` (Number): Day (1-31)

**Returns:** DateTime

**Examples:**
```
=DATE(2025, 10, 19)   // Returns October 19, 2025
=DATE(A1, A2, A3)     // Create date from cells
```

---

## File Functions

### FILE_SIZE
**Description:** Returns the size of a file in bytes.

**Syntax:** `=FILE_SIZE(file)`

**Parameters:**
- `file` (File or Text): File path

**Returns:** Number (bytes)

**Examples:**
```
=FILE_SIZE("C:\\documents\\report.pdf")  // Get file size
=FILE_SIZE(A1)                           // Size of file in A1
```

**Error Cases:** Returns error if file doesn't exist.

---

### FILE_EXTENSION
**Description:** Returns the file extension (e.g., ".txt", ".pdf").

**Syntax:** `=FILE_EXTENSION(file)`

**Parameters:**
- `file` (File or Text): File path

**Returns:** Text

**Examples:**
```
=FILE_EXTENSION("report.pdf")     // Returns ".pdf"
=FILE_EXTENSION(A1)               // Extension of file in A1
```

---

### FILE_NAME
**Description:** Returns the filename without the directory path.

**Syntax:** `=FILE_NAME(file)`

**Parameters:**
- `file` (File or Text): File path

**Returns:** Text

**Examples:**
```
=FILE_NAME("C:\\docs\\report.pdf")  // Returns "report.pdf"
=FILE_NAME(A1)                      // Name of file in A1
```

---

### FILE_READ
**Description:** Reads the contents of a text file.

**Syntax:** `=FILE_READ(file)`

**Parameters:**
- `file` (File or Text): File path

**Returns:** Text

**Examples:**
```
=FILE_READ("C:\\data\\config.txt")  // Read file contents
=FILE_READ(A1)                      // Read file from A1
```

**Error Cases:** Returns error if file doesn't exist or isn't readable.

---

## Directory Functions

### DIR_LIST
**Description:** Lists files in a directory.

**Syntax:** `=DIR_LIST(directory)`

**Parameters:**
- `directory` (Directory or Text): Directory path

**Returns:** Text (comma-separated filenames)

**Examples:**
```
=DIR_LIST("C:\\documents")    // List files
=DIR_LIST(A1)                 // List files from directory in A1
```

**Error Cases:** Returns error if directory doesn't exist.

---

### DIR_SIZE
**Description:** Calculates the total size of all files in a directory (including subdirectories).

**Syntax:** `=DIR_SIZE(directory)`

**Parameters:**
- `directory` (Directory or Text): Directory path

**Returns:** Number (total bytes)

**Examples:**
```
=DIR_SIZE("C:\\documents")    // Total size of directory
=DIR_SIZE(A1)                 // Size of directory in A1
```

**Error Cases:** Returns error if directory doesn't exist.

---

### DIRECTORY_TO_TABLE
**Description:** Expands a directory listing into a table with file information.

**Syntax:** `=DIRECTORY_TO_TABLE(directory)`

**Parameters:**
- `directory` (Directory or Text): Directory path

**Returns:** Table

**Examples:**
```
=DIRECTORY_TO_TABLE("C:\\documents")  // Create table of files
=DIRECTORY_TO_TABLE(A1)               // Table from directory in A1
```

---

## Table Functions

### TABLE_FILTER
**Description:** Filters rows in a table based on criteria.

**Syntax:** `=TABLE_FILTER(table, criteria)`

**Parameters:**
- `table` (Table): Table to filter
- `criteria` (Text): Filter criteria

**Returns:** Table

**Examples:**
```
=TABLE_FILTER(A1, "status=active")   // Filter by status
=TABLE_FILTER(A1, "value>100")       // Filter by numeric value
```

**Note:** Placeholder implementation - full filtering logic pending.

---

### TABLE_SORT
**Description:** Sorts a table by a specified column.

**Syntax:** `=TABLE_SORT(table, column)`

**Parameters:**
- `table` (Table): Table to sort
- `column` (Text): Column name to sort by

**Returns:** Table

**Examples:**
```
=TABLE_SORT(A1, "name")      // Sort by name column
=TABLE_SORT(A1, "date")      // Sort by date column
```

**Note:** Placeholder implementation - full sorting logic pending.

---

## Image Functions

### IMAGE_INFO
**Description:** Returns information about an image file (dimensions, format, size).

**Syntax:** `=IMAGE_INFO(image)`

**Parameters:**
- `image` (Image): Image file path

**Returns:** Text (image metadata)

**Examples:**
```
=IMAGE_INFO("C:\\photos\\image.jpg")  // Get image info
=IMAGE_INFO(A1)                       // Info for image in A1
```

**Error Cases:** Returns error if image file doesn't exist.

---

## PDF Functions

### PDF_TO_TEXT
**Description:** Extracts text content from a PDF file.

**Syntax:** `=PDF_TO_TEXT(pdf)`

**Parameters:**
- `pdf` (PDF): PDF file path

**Returns:** Text

**Examples:**
```
=PDF_TO_TEXT("C:\\docs\\report.pdf")  // Extract PDF text
=PDF_TO_TEXT(A1)                      // Extract from PDF in A1
```

**Note:** Placeholder implementation - requires PDF library integration.

---

### PDF_PAGE_COUNT
**Description:** Returns the number of pages in a PDF document.

**Syntax:** `=PDF_PAGE_COUNT(pdf)`

**Parameters:**
- `pdf` (PDF): PDF file path

**Returns:** Number

**Examples:**
```
=PDF_PAGE_COUNT("C:\\docs\\report.pdf")  // Count pages
=PDF_PAGE_COUNT(A1)                      // Pages in PDF from A1
```

**Note:** Placeholder implementation - requires PDF library integration.

---

## AI Functions

### IMAGE_TO_CAPTION
**Description:** Generates a descriptive caption for an image using AI vision models (GPT-4-Vision, LLaVA).

**Syntax:** `=IMAGE_TO_CAPTION(image)`

**Parameters:**
- `image` (Image): Image file to caption

**Returns:** Text (AI-generated caption)

**Examples:**
```
=IMAGE_TO_CAPTION("C:\\photos\\sunset.jpg")  // Caption an image
=IMAGE_TO_CAPTION(A1)                        // Caption image in A1
```

**Requirements:** Requires configured AI service (Azure OpenAI or Ollama).

---

### TEXT_TO_IMAGE
**Description:** Generates an image from a text prompt using AI (DALL-E 3).

**Syntax:** `=TEXT_TO_IMAGE(prompt)`

**Parameters:**
- `prompt` (Text): Text description of the image to generate

**Returns:** Image (path to generated image)

**Examples:**
```
=TEXT_TO_IMAGE("A serene mountain landscape at sunset")
=TEXT_TO_IMAGE(A1)  // Generate from prompt in A1
```

**Requirements:** Requires configured AI service (Azure OpenAI).

---

### TRANSLATE
**Description:** Translates text from one language to another using AI.

**Syntax:** `=TRANSLATE(text, target_language)`

**Parameters:**
- `text` (Text): Text to translate
- `target_language` (Text): Target language (e.g., "Spanish", "French", "Japanese")

**Returns:** Text (translated)

**Examples:**
```
=TRANSLATE("Hello", "Spanish")     // Returns "Hola"
=TRANSLATE(A1, "French")           // Translate A1 to French
=TRANSLATE("こんにちは", "English") // Translate Japanese to English
```

**Requirements:** Requires configured AI service.

---

### SUMMARIZE
**Description:** Creates a concise summary of a longer text using AI.

**Syntax:** `=SUMMARIZE(text, [max_words])`

**Parameters:**
- `text` (Text): Text to summarize
- `max_words` (Number, optional): Maximum summary length in words (default: 100)

**Returns:** Text (summary)

**Examples:**
```
=SUMMARIZE(A1)           // Summarize text in A1
=SUMMARIZE(A1, 50)       // Summarize to max 50 words
=SUMMARIZE(FILE_READ("article.txt"), 200)  // Summarize file content
```

**Requirements:** Requires configured AI service.

---

### CHAT
**Description:** Sends a message to an AI assistant and returns the response.

**Syntax:** `=CHAT(message, [system_prompt])`

**Parameters:**
- `message` (Text): Your message to the AI
- `system_prompt` (Text, optional): System instructions for the AI

**Returns:** Text (AI response)

**Examples:**
```
=CHAT("What is the capital of France?")
=CHAT("Explain photosynthesis", "You are a biology teacher")
=CHAT(A1, A2)  // Message and system prompt from cells
```

**Requirements:** Requires configured AI service.

---

### CODE_REVIEW
**Description:** Reviews code and provides suggestions for improvements, bugs, and best practices.

**Syntax:** `=CODE_REVIEW(code)`

**Parameters:**
- `code` (Code): Code to review (Python, C#, JavaScript, TypeScript, HTML, CSS)

**Returns:** Text (review comments)

**Examples:**
```
=CODE_REVIEW(A1)  // Review code in A1
=CODE_REVIEW(FILE_READ("script.py"))  // Review code from file
```

**Requirements:** Requires configured AI service.

---

### JSON_QUERY
**Description:** Queries JSON data using natural language and returns the result.

**Syntax:** `=JSON_QUERY(json, query)`

**Parameters:**
- `json` (JSON): JSON data to query
- `query` (Text): Natural language query

**Returns:** JSON (query result)

**Examples:**
```
=JSON_QUERY(A1, "Find all users with age > 30")
=JSON_QUERY(A1, "Get the total revenue")
=JSON_QUERY(FILE_READ("data.json"), "List all product names")
```

**Requirements:** Requires configured AI service.

---

### AI_EXTRACT
**Description:** Extracts specific information from text using AI (e.g., emails, dates, names, entities).

**Syntax:** `=AI_EXTRACT(text, extract_type)`

**Parameters:**
- `text` (Text): Text to analyze
- `extract_type` (Text): What to extract ("emails", "dates", "names", "entities")

**Returns:** Text (extracted information)

**Examples:**
```
=AI_EXTRACT(A1, "emails")    // Extract email addresses
=AI_EXTRACT(A1, "dates")     // Extract dates
=AI_EXTRACT(A1, "names")     // Extract person names
=AI_EXTRACT(A1, "entities")  // Extract all entities
```

**Requirements:** Requires configured AI service.

---

### SENTIMENT
**Description:** Analyzes the sentiment (positive/negative/neutral) of text using AI.

**Syntax:** `=SENTIMENT(text)`

**Parameters:**
- `text` (Text): Text to analyze

**Returns:** Text (sentiment analysis result)

**Examples:**
```
=SENTIMENT("I love this product!")           // Returns "Positive"
=SENTIMENT("This is terrible.")              // Returns "Negative"
=SENTIMENT(A1)                               // Analyze sentiment of A1
```

**Requirements:** Requires configured AI service.

---

## Function Categories

### Category Summary

| Category | Count | Description |
|----------|-------|-------------|
| Math     | 9     | Arithmetic and mathematical operations |
| Text     | 7     | String manipulation and formatting |
| DateTime | 3     | Date and time operations |
| File     | 4     | File system operations |
| Directory| 3     | Directory operations |
| Table    | 2     | Table manipulation (placeholders) |
| Image    | 1     | Image information |
| PDF      | 2     | PDF operations (placeholders) |
| AI       | 9     | AI-powered operations |
| **Total**| **40**| **All functions** |

---

## AI Service Configuration

AI functions require a configured AI service connection. Supported providers:

- **Azure OpenAI** (Recommended for production)
- **Ollama** (Local models)
- **OpenAI** (Direct API)

Configure AI services in **Settings → Service Connections**.

---

## Notes

- **Placeholder Functions:** TABLE_FILTER, TABLE_SORT, PDF_TO_TEXT, and PDF_PAGE_COUNT are placeholders awaiting full implementation.
- **AI Functions:** All AI functions require an active AI service connection and may incur API costs.
- **Error Handling:** Functions return Error type cells when operations fail (file not found, invalid input, etc.).
- **Async Execution:** All functions execute asynchronously and support cancellation.

---

## See Also

- [AiCalc User Guide](README.md)
- [Getting Started](QUICKSTART.md)
- [AI Service Setup](docs/AI_Services_Setup.md)
- [Formula Syntax](docs/Formula_Syntax.md)

---

**Document Version:** 1.0  
**Function Count:** 40  
**Last Verified:** October 19, 2025
