
# WordProcessor Utility

A lightweight and flexible Java library to dynamically populate Microsoft Word `.docx` templates with values from annotated Java POJO beans. This utility simplifies the generation of templated Word documents by leveraging annotations, reflection, and Apache POI under the hood.

---

## Features

- Populate Word documents using annotated POJO classes
- Supports checkbox replacements, text placeholders, and table rendering
- Chained API for clean, readable document generation
- Simple integration with any Java application

---

## Prerequisites

- Java 8 or higher
- Apache POI 5.4.4 (already integrated internally)

---

## Setup
To use this utility as a Maven dependency, add the following to your `pom.xml`:

```xml
<dependency>
    <groupId>com.github.salilvnair</groupId>
    <artifactId>wordprocessor</artifactId>
    <version>1.0.1</version>
</dependency>
```

---

## Usage Guide

### 1. Create a POJO

> Your POJO must **extend `BaseDocument`**

> Each field that should be replaced in the Word document must be annotated with `@PlaceHolder("key")`.

> To make key unique we can have it eclosed between `{{}}`, for example `{{key}}`

> The same has to be given in the `@Placeholder` `value` attribute

> For the `checkbox` type generally we define `tag value` in the docx and same name we will be giving in the placeholder value.

To mark a field as a checkbox, use the optional `checkbox = true` parameter:

```java
import com.github.salilvnair.wordprocessor.bean.BaseDocument;
import com.github.salilvnair.wordprocessor.reflect.annotation.PlaceHolder;

public class MyDocument extends BaseDocument {

    @PlaceHolder("{{name}}")
    private String name;

    @PlaceHolder("{{age}}")
    private String age;

    @PlaceHolder(value = "subscribed", checkbox = true)
    private boolean subscribed;

    // Getters and Setters
}
```

### 2. Prepare the `.docx` Template

Create a Microsoft Word `.docx` file and insert placeholders using the following format:

```
{{name}}
{{age}}
subscribed  <- Checkbox (in form-enabled documents)
```

> 🔐 Ensure your `.docx`  template is **not password protected** and is in a valid editable format.

> ✅ For Checkboxes: Use `@PlaceHolder(checkbox = true)` on the field to treat that as the checkbox.

---

### 3. Call the Utility

```java
import com.github.salilvnair.wordprocessor.writer.WordDocumentWriter;

public class WordProcessorRunner {
    public static void main(String[] args) {
        MyDocument doc = new MyDocument();
        doc.setName("Alice");
        doc.setAge("30");
        doc.setSubscribed(true);

        WordDocumentWriter writer = new WordDocumentWriter();
        try {
            writer.placeHolderBeans(doc)
                  .template("/path/to/template.docx")
                  .table() // Optional - use if template has table placeholders otherwise use .text()
                  .replace()
                  .save("/path/to/output.docx");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
```

---

## API Reference

### `WordDocumentWriter`

- `.placeHolderBeans(Object... beans)`: Accepts one or more POJO objects with annotated fields.
- `.template(String path)`: Path to your `.docx` template.
- `.table()`: Enables table processing (optional).
- `.replace()`: Triggers the replacement of placeholders.
- `.save(String outputPath)`: Saves the generated document to the given path.

---

## Examples

See the `test/main/TestWordProcessor.java` and `test/placeholder/bean/TestBean.java` for real-world usage examples.

---

## Contributing

Feel free to fork, enhance, or submit issues or pull requests on GitHub.

---

## License

MIT License
