# BroggiSoft.OfficeExport

Generate documents with the BroggiSoft.OfficeExport package. Simply use the tag system to transform your templates into final Word, Excel or PowerPoint documents.
The package will retain the font, size, and other formatting, and can generate rows in a table if the object is an array, or replace an existing rectangle shape with your image (your logo for example).
Leave the management of documents (contracts, letters, etc.) to the business users, while developers only need to supply the data model in the form of a dictionary, JSON, or object to seamlessly populate your Word, Excel, or PowerPoint documents.

### Features

- Replace simple text tags in Word, Excel and PowerPoint templates with JSON, dictionary or object types.
- Support for arrays, you can send an array of objects, of string, of int... and it will add as many rows in a Word, Excel or PowerPoint table. 
- Support for images, you can add an image from a base64 format.

### How to use

- Create a template (Word, Excel or PowerPoint) with tags, for example {{city}}
- Create a JSON, dictionary or object with the corresponding values
- Use BroggiSoft.OfficeExport to replace the tags with the values and generate the final document
- For an array, create a table in your template et set a tag in a row like {{YourArray[i].YourProperty}} or {{YourArray[2].YourProperty}} for a specific position. This will add as many rows as there are in the array.  {{Product[i].Price}} : for each product a row will be added in the table and set the value Price of this product.
- For add an image, insert a rectangle shape with the excpected size int your template, and set in this shape a text tag {{IMG=YourNameTag}}, the value of your image must be in base64 format.

### Examples


            string jsonValues = "{\"name\":\"John\",\"city\":\"New York\"}";

            OfficeExport.ExportFromJson("WordTemplate.docx", "WordResult.docx", jsonValues);

    //Or with an object:

            ExampleModel example = new ExampleModel();
            example.value1 = "Hello";
            example.value2 = "World";

            OfficeExport.ExportFromObject("WordTemplate.docx", "WordResult.docx", example);

### Use

This tool is completely free. All the necessary documentation can be found at: https://www.broggisoft.com/wiki

Terms and conditions can be found here: https://www.broggisoft.com/termsandconditions

If you have any questions, please don't hesitate to contact me at contact@broggisoft.com.