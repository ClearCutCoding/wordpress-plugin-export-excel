# Export Excel Plugin

Allows for exporting of data tables into Excel.

### Version

1.0.0

---

## Description

This plugin allows you to create data tables in a post or page and export these tables to an excel spreadsheet.

Each table that you want to export will be converted to a separate worksheet within Excel.

## Installation

1. Move 'export-excel' into the wordpress plugins folder.
2. Activate the plugin through the backend.
3. Make sure you have data tables in your post.  These can be entered manually or using another plugin.
4. For any tables you want to export, add an attribute 'export' with the value 'yes', such as: 
```
<table export="yes">
```
5. To display the export button, add the following shortcode to the content wherever you wish the button to appear: [exportexcel]

## Resources

### Libraries Used

* PHPExcel
