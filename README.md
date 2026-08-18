![Aspose.Cells Cloud SDK for Perl](https://img.shields.io/badge/aspose.cells%20Cloud%20SDK-26.8-green?style=for-the-badge&logo=perl) [![Product Page](https://img.shields.io/badge/Product-0288d1?style=for-the-badge&logo=Google-Chrome&logoColor=white)](https://products.aspose.cloud/cells/perl/) [![Documentation](https://img.shields.io/badge/Documentation-388e3c?style=for-the-badge&logo=Hugo&logoColor=white)](https://docs.aspose.cloud/cells/) [![API Ref](https://img.shields.io/badge/Reference-f39c12?style=for-the-badge&logo=html5&logoColor=white)](https://reference.aspose.cloud/cells/) [![Examples](https://img.shields.io/badge/Examples-1565c0?style=for-the-badge&logo=Github&logoColor=white)](https://github.com/aspose-cells-cloud/aspose-cells-cloud-perl/tree/master/examples) [![Blog](https://img.shields.io/badge/Blog-d32f2f?style=for-the-badge&logo=WordPress&logoColor=white)](https://blog.aspose.cloud/categories/aspose.cells-cloud-product-family/) [![Support](https://img.shields.io/badge/Support-7b1fa2?style=for-the-badge&logo=Discourse&logoColor=white)](https://forum.aspose.cloud/c/cells/7) [![License](https://img.shields.io/badge/License-7b1fa2?style=for-the-badge&logo=Discourse&logoColor=white)](https://forum.aspose.cloud/c/cells/7) [![License](https://img.shields.io/github/license/aspose-cells-cloud/aspose-cells-cloud-go?style=for-the-badge&logo=rocket&logoColor=white)](https://github.com/aspose-cells-cloud/aspose-cells-cloud-go/blob/master/LICENSE) ![CPAN](https://img.shields.io/cpan/v/AsposeCellsCloud-CellsApi?style=for-the-badge&logo=rocket&logoColor=white)

<p align="center">
  <a href="#english">English</a> |
  <a href="#中文">中文</a> |
  <a href="#日本語">日本語</a> |
  <a href="#deutsch">Deutsch</a>
</p>

---

<h1 id="english">English</h1>

[Aspose.Cells Cloud SDK for Perl](https://products.aspose.cloud/cells/perl) is a cloud-native REST API that enables Perl developers to **create**, **read**, **edit**, **convert**, and **repair** spreadsheet files—including **Excel** (**XLS**, **XLSX**, **XLSB**, **XLSM**), **OpenDocument Spreadsheet** (**ODS**), **CSV**, **TSV**, **JSON**, **HTML**, **PDF**, and **more—without requiring Microsoft Excel or Office to be installed**.

Built on the **Aspose.Cells Cloud Web API**, this MIT-licensed SDK supports advanced spreadsheet operations such as:

- Cell formatting, formulas, and data validation
- Pivot tables, charts, hyperlinks, and comments
- Conditional formatting and smart markers
- Worksheet merging, splitting, and protection
- Batch processing and background removal

It seamlessly integrates with **AWS**, **Microsoft Azure**, and **Google Cloud**, ensuring **high availability**, **scalability**, and **data integrity**. Ideal for serverless apps, microservices, and cloud automation workflows.

## Quick Start Guide

To begin with Aspose.Cells Cloud, here is what you need to do:

1. Sign up for an account at [Aspose for Cloud](https://dashboard.aspose.cloud/#/apps) to obtain your application details.
2. Install the Aspose.Cells Cloud Perl module from the [CPAN distribution](https://www.cpan.org/).
3. Use the conversion code provided below as a reference to add or modify your application.

## Convert an Excel File Using Perl

```perl
use strict;
use warnings;
use File::Slurp;
use MIME::Base64;
use AsposeCellsCloud::ApiClient;
use AsposeCellsCloud::CellsApi;
use AsposeCellsCloud::Configuration;
use AsposeCellsCloud::Request::ConvertSpreadsheetRequest;

my $config = AsposeCellsCloud::Configuration->new( client_id => $ENV{'CellsCloudClientId'}, client_secret => $ENV{'CellsCloudClientSecret'});
my $instance = AsposeCellsCloud::CellsApi->new(AsposeCellsCloud::ApiClient->new( $config));

my $request = AsposeCellsCloud::Request::ConvertSpreadsheetRequest->new();
$request->{spreadsheet} =  'EmployeeSalesSummary.xlsx';
$request->{format} = 'pdf';
my $response = $instance->put_convert_workbook(request=> $request);
open (my $fh, '>', 'EmployeeSalesSummary.pdf') or die "Cannot open EmployeeSalesSummary.pdf $!";
binmode($fh);
print $fh $response;
close($fh);
```

## Supported File Formats

| **Format**                                                        | **Description**                                                                                                                                                                  | **Load** | **Save** |
| :---------------------------------------------------------------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | :------- | :------- |
| [XLS](https://docs.fileformat.com/spreadsheet/xls/)               | Excel 95/5.0 - 2003 Workbook.                                                                                                                                                    | &radic;  | &radic;  |
| [XLSX](https://docs.fileformat.com/spreadsheet/xlsx/)             | Office Open XML SpreadsheetML Workbook or template file, with or without macros.                                                                                                 | &radic;  | &radic;  |
| [XLSB](https://docs.fileformat.com/spreadsheet/xlsb/)             | Excel Binary Workbook.                                                                                                                                                           | &radic;  | &radic;  |
| [XLSM](https://docs.fileformat.com/spreadsheet/xlsm/)             | Excel Macro-Enabled Workbook.                                                                                                                                                    | &radic;  | &radic;  |
| [XLT](https://docs.fileformat.com/spreadsheet/xlt/)               | Excel 97 - Excel 2003 Template.                                                                                                                                                  | &radic;  | &radic;  |
| [XLTX](https://docs.fileformat.com/spreadsheet/xltx/)             | Excel Template.                                                                                                                                                                  | &radic;  | &radic;  |
| [XLTM](https://docs.fileformat.com/spreadsheet/xltm/)             | Excel Macro-Enabled Template.                                                                                                                                                    | &radic;  | &radic;  |
| [XLAM](https://docs.fileformat.com/spreadsheet/xlam/)             | An Excel Macro-Enabled Add-In file used to add new functions to Excel.                                                                                                           |          | &radic;  |
| [CSV](https://docs.fileformat.com/spreadsheet/csv/)               | CSV (Comma Separated Value) file.                                                                                                                                                | &radic;  | &radic;  |
| [TSV](https://docs.fileformat.com/spreadsheet/tsv/)               | TSV (Tab-separated values) file.                                                                                                                                                 | &radic;  | &radic;  |
| [TXT](https://docs.fileformat.com/word-processing/txt/)           | Delimited plain text file.                                                                                                                                                       | &radic;  | &radic;  |
| [HTML](https://docs.fileformat.com/web/html/)                     | HTML format.                                                                                                                                                                     | &radic;  | &radic;  |
| [MHTML](https://docs.fileformat.com/web/mhtml/)                   | MHTML file.                                                                                                                                                                      | &radic;  | &radic;  |
| [ODS](https://docs.fileformat.com/spreadsheet/ods/)               | ODS (OpenDocument Spreadsheet).                                                                                                                                                  | &radic;  | &radic;  |
| [Numbers](https://docs.fileformat.com/spreadsheet/numbers/)       | Documents created by Apple's "Numbers" application, which is part of Apple's iWork office suite running on Mac OS X and iOS operating systems.                                   | &radic;  |          |
| [JSON](https://docs.fileformat.com/web/json/)                     | JavaScript Object Notation.                                                                                                                                                      | &radic;  | &radic;  |
| [DIF](https://docs.fileformat.com/spreadsheet/dif/)               | Data Interchange Format.                                                                                                                                                         |          | &radic;  |
| [PDF](https://docs.fileformat.com/pdf/)                           | Adobe Portable Document Format.                                                                                                                                                  |          | &radic;  |
| [XPS](https://docs.fileformat.com/page-description-language/xps/) | XML Paper Specification Format.                                                                                                                                                  |          | &radic;  |
| [SVG](https://docs.fileformat.com/page-description-language/svg/) | Scalable Vector Graphics Format.                                                                                                                                                 |          | &radic;  |
| [TIFF](https://docs.fileformat.com/image/tiff/)                   | Tagged Image File Format.                                                                                                                                                        |          | &radic;  |
| [PNG](https://docs.fileformat.com/image/png/)                     | Portable Network Graphics Format.                                                                                                                                                |          | &radic;  |
| [BMP](https://docs.fileformat.com/image/bmp/)                     | Bitmap Image Format.                                                                                                                                                             |          | &radic;  |
| [EMF](https://docs.fileformat.com/image/emf/)                     | Enhanced Metafile Format.                                                                                                                                                        |          | &radic;  |
| [JPEG](https://docs.fileformat.com/image/jpeg/)                   | JPEG is an image format that uses lossy compression.                                                                                                                             |          | &radic;  |
| [GIF](https://docs.fileformat.com/image/gif/)                     | Graphical Interchange Format.                                                                                                                                                    |          | &radic;  |
| [MARKDOWN](https://docs.fileformat.com/word-processing/md/)       | Represents a Markdown document.                                                                                                                                                  |          | &radic;  |
| [SXC](https://docs.fileformat.com/spreadsheet/sxc/)               | An XML-based format used by OpenOffice and StarOffice.                                                                                                                           | &radic;  | &radic;  |
| [FODS](https://docs.fileformat.com/spreadsheet/fods/)             | An Open Document format stored as flat XML.                                                                                                                                      | &radic;  | &radic;  |
| [DOCX](https://docs.fileformat.com/word-processing/docx/)         | A well-known format for Microsoft Word documents combining XML and binary files.                                                                                                 |          | &radic;  |
| [PPTX](https://docs.fileformat.com/presentation/pptx/)            | The PPTX format is based on the Microsoft PowerPoint Open XML presentation file format.                                                                                          |          | &radic;  |
| [OTS](https://docs.fileformat.com/spreadsheet/ots/)               | OTS (OpenDocument Spreadsheet).                                                                                                                                                  | &radic;  | &radic;  |
| [XML](https://docs.fileformat.com/web/xml/)                       | XML file.                                                                                                                                                                        | &radic;  | &radic;  |
| [HTM](https://docs.fileformat.com/web/htm/)                       | HTM file.                                                                                                                                                                        | &radic;  | &radic;  |
| [TIF](https://docs.fileformat.com/image/tiff/)                    | Tagged Image File Format.                                                                                                                                                        |          | &radic;  |
| [WMF](https://docs.fileformat.com/image/wmf/)                     | Windows Metafile Format.                                                                                                                                                         |          | &radic;  |
| [PCL](https://docs.fileformat.com/page-description-language/pcl/) | Printer Command Language Format.                                                                                                                                                 |          | &radic;  |
| [AZW3](https://docs.fileformat.com/ebook/azw3/)                   | AZW3/KF8 File Format.                                                                                                                                                            |          | &radic;  |
| [EPUB](https://docs.fileformat.com/ebook/epub/)                   | EPUB File Format.                                                                                                                                                                |          | &radic;  |
| [DBF](https://docs.fileformat.com/database/dbf/)                  | Database File Format.                                                                                                                                                            |          | &radic;  |
| [XHTML](https://docs.fileformat.com/web/xhtml/)                   | XHTML File Format.                                                                                                                                                               |          | &radic;  |

## Architecture

![architecture.png](architecture.png)

## [Developer Reference](docs/DeveloperGuide.md#overview)

### Manipulate Excel and Other Spreadsheet Files in the Cloud

- **File Manipulation**: Users can upload, download, delete, and manage Excel files stored in the cloud.
- **File Formatting**: Supports formatting of cells, fonts, colors, and alignment modes in Excel files to cater to specific requirements.
- **Data Processing**: Powerful functions for data processing including reading, writing, modifying cell data, performing formula calculations, and formatting data.
- **Formula Calculation**: Built-in formula engine handles complex formula calculations in Excel and returns accurate results.
- **Chart Manipulation**: Users can create, edit, and delete charts from Excel files for data analysis and visualization needs.
- **Table Processing**: Offers robust processing capabilities for various format operations such as creation, editing, formatting, and conversion, meeting diverse format processing needs.
- **Data Validation**: Includes data validation functions to set cell data type, range, and uniqueness, ensuring data accuracy and integrity.
- **Batch Processing**: Supports batch processing of multiple Excel documents, such as batch format conversion, data extraction, and style application.
- **Import/Export Data**: Facilitates importing data from various sources into spreadsheets and exporting spreadsheet data to other formats.
- **Security Management**: Offers a range of security features like data encryption, access control, and permission management to safeguard the security and integrity of spreadsheet data.

## Features & Enhancements in Version 26.8

| **Summary** | **Category** |
| :--- | :--- |
| Enchent smart template feature. | Improvement |
| A new AI data analysis API has been added. | New Feature |

## Available SDKs

The Aspose.Cells Cloud SDK is available in multiple popular programming languages, enabling developers to integrate spreadsheet processing capabilities across various development environments.

[![Go](https://img.shields.io/badge/Go-00ADD8.svg?style=for-the-badge&logo=go&logoColor=white)](https://github.com/aspose-cells-cloud/aspose-cells-cloud-go) [![Go](https://img.shields.io/badge/Go-Install%20go%20get%20package--asposecellscloud-%2300ADD8?logo=go&style=for-the-badge)](https://pkg.go.dev/github.com/aspose-cells-cloud/aspose-cells-cloud-go/v25)

[![Java](https://img.shields.io/badge/Java-red?logo=openjdk&style=for-the-badge&logoColor=white)](https://github.com/aspose-cells-cloud/aspose-cells-cloud-java) [![Java](https://img.shields.io/badge/Maven-Aspose.Cells%20Cloud.pom.xml-red?logo=apache-maven&style=for-the-badge)](https://github.com/aspose-cells-cloud/aspose-cells-cloud-java/blob/master/Aspose.Cells.Cloud.pom.xml)

[![.NET](https://img.shields.io/badge/.NET-%23512BD4?style=for-the-badge&logo=dotnet&logoColor=white)](https://github.com/aspose-cells-cloud/aspose-cells-cloud-dotnet) [![.NET](https://img.shields.io/badge/NuGet-Install%20Aspose.Cells--Cloud-%23512BD4?logo=nuget&style=for-the-badge)](https://www.nuget.org/packages/Aspose.cells-Cloud/#readme-body-tab)

[![Node.js](https://img.shields.io/badge/Node.js-43853D.svg?style=for-the-badge&logo=node.js&logoColor=white)](https://github.com/aspose-cells-cloud/aspose-cells-cloud-node) [![Node.js](https://img.shields.io/badge/npm-install%20asposecellscloud-orange?logo=npm&style=for-the-badge)](https://www.npmjs.com/package/asposecellscloud)

[![Perl](https://img.shields.io/badge/Perl-39457E.svg?style=for-the-badge&logo=perl&logoColor=white)](https://github.com/aspose-cells-cloud/aspose-cells-cloud-perl) [![Perl](https://img.shields.io/badge/CPAN-Install%20AsposeCellsCloud--CellsApi-blue?logo=perl&style=for-the-badge)](https://metacpan.org/dist/AsposeCellsCloud-CellsApi)

[![PHP](https://img.shields.io/badge/PHP-777BB4.svg?style=for-the-badge&logo=php&logoColor=white)](https://github.com/aspose-cells-cloud/aspose-cells-cloud-php) [![PHP](https://img.shields.io/badge/Composer-require%20aspose/cells--sdk--php-8892BF?logo=php&style=for-the-badge)](https://packagist.org/packages/aspose/cells-sdk-php)

[![Python](https://img.shields.io/badge/Python-14354C.svg?style=for-the-badge&logo=python&logoColor=white)](https://github.com/aspose-cells-cloud/aspose-cells-cloud-python) [![Python](https://img.shields.io/badge/pip-install%20asposecellscloud-blue?logo=pypi&style=for-the-badge)](https://pypi.org/project/asposecellscloud/)

[![Ruby](https://img.shields.io/badge/Ruby-CC342D.svg?style=for-the-badge&logo=ruby&logoColor=white)](https://github.com/aspose-cells-cloud/aspose-cells-cloud-ruby) [![Ruby](https://img.shields.io/badge/Gem-install%20aspose__cells__cloud-red?logo=ruby&style=for-the-badge)](https://rubygems.org/gems/aspose_cells_cloud)

## [Release History](CHANGELOG.md)

---

<h1 id="中文">中文</h1>

[Aspose.Cells Cloud SDK for Perl](https://products.aspose.cloud/cells/perl) 是一个云原生 REST API，使 Perl 开发者能够**创建**、**读取**、**编辑**、**转换**和**修复**电子表格文件——包括 **Excel**（**XLS**、**XLSX**、**XLSB**、**XLSM**）、**OpenDocument Spreadsheet**（**ODS**）、**CSV**、**TSV**、**JSON**、**HTML**、**PDF** 等——**无需安装 Microsoft Excel 或 Office**。

基于 **Aspose.Cells Cloud Web API** 构建，此 MIT 许可的 SDK 支持高级电子表格操作，例如：

- 单元格格式化、公式和数据验证
- 数据透视表、图表、超链接和注释
- 条件格式和智能标记
- 工作表合并、拆分和保护
- 批量处理和背景移除

它无缝集成 **AWS**、**Microsoft Azure** 和 **Google Cloud**，确保**高可用性**、**可扩展性**和**数据完整性**。非常适合无服务器应用、微服务和云自动化工作流。

## 快速入门指南

要开始使用 Aspose.Cells Cloud，您需要执行以下操作：

1. 在 [Aspose for Cloud](https://dashboard.aspose.cloud/#/apps) 注册账户以获取您的应用程序凭据。
2. 从 [CPAN 发行版](https://www.cpan.org/) 安装 Aspose.Cells Cloud Perl 模块。
3. 使用下面提供的转换代码作为参考，添加或修改您的应用程序。

## 使用 Perl 转换 Excel 文件

```perl
use strict;
use warnings;
use File::Slurp;
use MIME::Base64;
use AsposeCellsCloud::ApiClient;
use AsposeCellsCloud::CellsApi;
use AsposeCellsCloud::Configuration;
use AsposeCellsCloud::Request::ConvertSpreadsheetRequest;

my $config = AsposeCellsCloud::Configuration->new( client_id => $ENV{'CellsCloudClientId'}, client_secret => $ENV{'CellsCloudClientSecret'});
my $instance = AsposeCellsCloud::CellsApi->new(AsposeCellsCloud::ApiClient->new( $config));

my $request = AsposeCellsCloud::Request::ConvertSpreadsheetRequest->new();
$request->{spreadsheet} =  'EmployeeSalesSummary.xlsx';
$request->{format} = 'pdf';
my $response = $instance->put_convert_workbook(request=> $request);
open (my $fh, '>', 'EmployeeSalesSummary.pdf') or die "Cannot open EmployeeSalesSummary.pdf $!";
binmode($fh);
print $fh $response;
close($fh);
```

## 支持的文件格式

| **格式**                                                          | **描述**                                                                                                                      | **加载** | **保存** |
| :---------------------------------------------------------------- | :---------------------------------------------------------------------------------------------------------------------------- | :------- | :------- |
| [XLS](https://docs.fileformat.com/spreadsheet/xls/)               | Excel 95/5.0 - 2003 工作簿。                                                                                                  | &radic;  | &radic;  |
| [XLSX](https://docs.fileformat.com/spreadsheet/xlsx/)             | Office Open XML SpreadsheetML 工作簿或模板文件，可包含宏。                                                                     | &radic;  | &radic;  |
| [XLSB](https://docs.fileformat.com/spreadsheet/xlsb/)             | Excel 二进制工作簿。                                                                                                          | &radic;  | &radic;  |
| [XLSM](https://docs.fileformat.com/spreadsheet/xlsm/)             | Excel 启用宏的工作簿。                                                                                                        | &radic;  | &radic;  |
| [XLT](https://docs.fileformat.com/spreadsheet/xlt/)               | Excel 97 - Excel 2003 模板。                                                                                                  | &radic;  | &radic;  |
| [XLTX](https://docs.fileformat.com/spreadsheet/xltx/)             | Excel 模板。                                                                                                                  | &radic;  | &radic;  |
| [XLTM](https://docs.fileformat.com/spreadsheet/xltm/)             | Excel 启用宏的模板。                                                                                                          | &radic;  | &radic;  |
| [XLAM](https://docs.fileformat.com/spreadsheet/xlam/)             | 用于向 Excel 添加新功能的 Excel 启用宏的加载项文件。                                                                           |          | &radic;  |
| [CSV](https://docs.fileformat.com/spreadsheet/csv/)               | CSV（逗号分隔值）文件。                                                                                                       | &radic;  | &radic;  |
| [TSV](https://docs.fileformat.com/spreadsheet/tsv/)               | TSV（制表符分隔值）文件。                                                                                                     | &radic;  | &radic;  |
| [TXT](https://docs.fileformat.com/word-processing/txt/)           | 带分隔符的纯文本文件。                                                                                                        | &radic;  | &radic;  |
| [HTML](https://docs.fileformat.com/web/html/)                     | HTML 格式。                                                                                                                   | &radic;  | &radic;  |
| [MHTML](https://docs.fileformat.com/web/mhtml/)                   | MHTML 文件。                                                                                                                  | &radic;  | &radic;  |
| [ODS](https://docs.fileformat.com/spreadsheet/ods/)               | ODS（OpenDocument 电子表格）。                                                                                                 | &radic;  | &radic;  |
| [Numbers](https://docs.fileformat.com/spreadsheet/numbers/)       | 由 Apple "Numbers" 应用程序创建的文档，该应用程序是 Apple iWork 办公套件的一部分，运行在 Mac OS X 和 iOS 操作系统上。            | &radic;  |          |
| [JSON](https://docs.fileformat.com/web/json/)                     | JavaScript 对象表示法。                                                                                                      | &radic;  | &radic;  |
| [DIF](https://docs.fileformat.com/spreadsheet/dif/)               | 数据交换格式。                                                                                                                |          | &radic;  |
| [PDF](https://docs.fileformat.com/pdf/)                           | Adobe 便携文档格式。                                                                                                           |          | &radic;  |
| [XPS](https://docs.fileformat.com/page-description-language/xps/) | XML 纸张规范格式。                                                                                                             |          | &radic;  |
| [SVG](https://docs.fileformat.com/page-description-language/svg/) | 可缩放矢量图形格式。                                                                                                          |          | &radic;  |
| [TIFF](https://docs.fileformat.com/image/tiff/)                   | 标记图像文件格式。                                                                                                            |          | &radic;  |
| [PNG](https://docs.fileformat.com/image/png/)                     | 便携网络图形格式。                                                                                                            |          | &radic;  |
| [BMP](https://docs.fileformat.com/image/bmp/)                     | 位图图像格式。                                                                                                                |          | &radic;  |
| [EMF](https://docs.fileformat.com/image/emf/)                     | 增强型图元文件格式。                                                                                                          |          | &radic;  |
| [JPEG](https://docs.fileformat.com/image/jpeg/)                   | JPEG 是一种使用有损压缩的图像格式。                                                                                           |          | &radic;  |
| [GIF](https://docs.fileformat.com/image/gif/)                     | 图形交换格式。                                                                                                                |          | &radic;  |
| [MARKDOWN](https://docs.fileformat.com/word-processing/md/)       | 表示 Markdown 文档。                                                                                                          |          | &radic;  |
| [SXC](https://docs.fileformat.com/spreadsheet/sxc/)               | OpenOffice 和 StarOffice 使用的基于 XML 的格式。                                                                            | &radic;  | &radic;  |
| [FODS](https://docs.fileformat.com/spreadsheet/fods/)             | 以平面 XML 存储的 Open Document 格式。                                                                                      | &radic;  | &radic;  |
| [DOCX](https://docs.fileformat.com/word-processing/docx/)         | 广为人知的 Microsoft Word 文档格式，结合了 XML 和二进制文件。                                                               |          | &radic;  |
| [PPTX](https://docs.fileformat.com/presentation/pptx/)            | 基于 Microsoft PowerPoint Open XML 演示文稿文件格式。                                                                          |          | &radic;  |
| [OTS](https://docs.fileformat.com/spreadsheet/ots/)               | OTS（OpenDocument 电子表格）。                                                                                                 | &radic;  | &radic;  |
| [XML](https://docs.fileformat.com/web/xml/)                       | XML 文件。                                                                                                                    | &radic;  | &radic;  |
| [HTM](https://docs.fileformat.com/web/htm/)                       | HTM 文件。                                                                                                                    | &radic;  | &radic;  |
| [TIF](https://docs.fileformat.com/image/tiff/)                    | 标记图像文件格式。                                                                                                            |          | &radic;  |
| [WMF](https://docs.fileformat.com/image/wmf/)                     | Windows 图元文件格式。                                                                                                        |          | &radic;  |
| [PCL](https://docs.fileformat.com/page-description-language/pcl/) | 打印机命令语言格式。                                                                                                          |          | &radic;  |
| [AZW3](https://docs.fileformat.com/ebook/azw3/)                   | AZW3/KF8 文件格式。                                                                                                           |          | &radic;  |
| [EPUB](https://docs.fileformat.com/ebook/epub/)                   | EPUB 文件格式。                                                                                                               |          | &radic;  |
| [DBF](https://docs.fileformat.com/database/dbf/)                  | 数据库文件格式。                                                                                                              |          | &radic;  |
| [XHTML](https://docs.fileformat.com/web/xhtml/)                   | XHTML 文件格式。                                                                                                              |          | &radic;  |

## 架构

![architecture.png](architecture.png)

## [开发者参考](docs/DeveloperGuide.md#overview)

### 在云中操作 Excel 及其他电子表格文件

- **文件操作**：用户可以上传、下载、删除和管理存储在云端的 Excel 文件。
- **文件格式化**：支持 Excel 文件中单元格、字体、颜色和对齐方式的格式化，以满足特定需求。
- **数据处理**：强大的数据处理功能，包括读取、写入、修改单元格数据、执行公式计算和格式化数据。
- **公式计算**：内置公式引擎可处理 Excel 中的复杂公式计算并返回准确结果。
- **图表操作**：用户可以创建、编辑和删除 Excel 文件中的图表，满足数据分析和可视化需求。
- **表格处理**：提供强大的格式操作处理能力，包括创建、编辑、格式化和转换等功能。
- **数据验证**：包含数据验证功能，可设置单元格数据类型、范围和唯一性，确保数据准确性和完整性。
- **批量处理**：支持批量处理多个 Excel 文档，如批量格式转换、数据提取和样式应用。
- **数据导入/导出**：支持从各种数据源导入数据到电子表格，以及将电子表格数据导出到其他格式。
- **安全管理**：提供数据加密、访问控制和权限管理等一系列安全功能，保障电子表格数据的安全性和完整性。

## 版本 26.8 的功能与增强

| **摘要** | **类别** |
| :--- | :--- |
| Enchent smart template feature. | Improvement |
| A new AI data analysis API has been added. | New Feature |

## 可用的 SDK

Aspose.Cells Cloud SDK 提供多种流行编程语言版本，使开发者能够在各种开发环境中集成电子表格处理功能。

[![Go](https://img.shields.io/badge/Go-00ADD8.svg?style=for-the-badge&logo=go&logoColor=white)](https://github.com/aspose-cells-cloud/aspose-cells-cloud-go) [![Go](https://img.shields.io/badge/Go-Install%20go%20get%20package--asposecellscloud-%2300ADD8?logo=go&style=for-the-badge)](https://pkg.go.dev/github.com/aspose-cells-cloud/aspose-cells-cloud-go/v25)

[![Java](https://img.shields.io/badge/Java-red?logo=openjdk&style=for-the-badge&logoColor=white)](https://github.com/aspose-cells-cloud/aspose-cells-cloud-java) [![Java](https://img.shields.io/badge/Maven-Aspose.Cells%20Cloud.pom.xml-red?logo=apache-maven&style=for-the-badge)](https://github.com/aspose-cells-cloud/aspose-cells-cloud-java/blob/master/Aspose.Cells.Cloud.pom.xml)

[![.NET](https://img.shields.io/badge/.NET-%23512BD4?style=for-the-badge&logo=dotnet&logoColor=white)](https://github.com/aspose-cells-cloud/aspose-cells-cloud-dotnet) [![.NET](https://img.shields.io/badge/NuGet-Install%20Aspose.Cells--Cloud-%23512BD4?logo=nuget&style=for-the-badge)](https://www.nuget.org/packages/Aspose.cells-Cloud/#readme-body-tab)

[![Node.js](https://img.shields.io/badge/Node.js-43853D.svg?style=for-the-badge&logo=node.js&logoColor=white)](https://github.com/aspose-cells-cloud/aspose-cells-cloud-node) [![Node.js](https://img.shields.io/badge/npm-install%20asposecellscloud-orange?logo=npm&style=for-the-badge)](https://www.npmjs.com/package/asposecellscloud)

[![Perl](https://img.shields.io/badge/Perl-39457E.svg?style=for-the-badge&logo=perl&logoColor=white)](https://github.com/aspose-cells-cloud/aspose-cells-cloud-perl) [![Perl](https://img.shields.io/badge/CPAN-Install%20AsposeCellsCloud--CellsApi-blue?logo=perl&style=for-the-badge)](https://metacpan.org/dist/AsposeCellsCloud-CellsApi)

[![PHP](https://img.shields.io/badge/PHP-777BB4.svg?style=for-the-badge&logo=php&logoColor=white)](https://github.com/aspose-cells-cloud/aspose-cells-cloud-php) [![PHP](https://img.shields.io/badge/Composer-require%20aspose/cells--sdk--php-8892BF?logo=php&style=for-the-badge)](https://packagist.org/packages/aspose/cells-sdk-php)

[![Python](https://img.shields.io/badge/Python-14354C.svg?style=for-the-badge&logo=python&logoColor=white)](https://github.com/aspose-cells-cloud/aspose-cells-cloud-python) [![Python](https://img.shields.io/badge/pip-install%20asposecellscloud-blue?logo=pypi&style=for-the-badge)](https://pypi.org/project/asposecellscloud/)

[![Ruby](https://img.shields.io/badge/Ruby-CC342D.svg?style=for-the-badge&logo=ruby&logoColor=white)](https://github.com/aspose-cells-cloud/aspose-cells-cloud-ruby) [![Ruby](https://img.shields.io/badge/Gem-install%20aspose__cells__cloud-red?logo=ruby&style=for-the-badge)](https://rubygems.org/gems/aspose_cells_cloud)

## [发布历史](CHANGELOG.md)

---

<h1 id="日本語">日本語</h1>

[Aspose.Cells Cloud SDK for Perl](https://products.aspose.cloud/cells/perl) は、Perl 開発者が **Excel**（**XLS**、**XLSX**、**XLSB**、**XLSM**）、**OpenDocument Spreadsheet**（**ODS**）、**CSV**、**TSV**、**JSON**、**HTML**、**PDF** などのスプレッドシートファイルを **Microsoft Excel や Office をインストールせずに** **作成**、**読み取り**、**編集**、**変換**、**修復** できるクラウドネイティブ REST API です。

**Aspose.Cells Cloud Web API** 上に構築されたこの MIT ライセンスの SDK は、以下のような高度なスプレッドシート操作をサポートします：

- セルの書式設定、数式、データ検証
- ピボットテーブル、グラフ、ハイパーリンク、コメント
- 条件付き書式とスマートマーカー
- ワークシートの結合、分割、保護
- バッチ処理と背景除去

**AWS**、**Microsoft Azure**、**Google Cloud** とシームレスに統合し、**高可用性**、**スケーラビリティ**、**データ整合性** を確保します。サーバーレスアプリ、マイクロサービス、クラウド自動化ワークフローに最適です。

## クイックスタートガイド

Aspose.Cells Cloud を始めるには、以下の手順を実行してください：

1. [Aspose for Cloud](https://dashboard.aspose.cloud/#/apps) でアカウントを作成し、アプリケーションの詳細情報を取得します。
2. [CPAN ディストリビューション](https://www.cpan.org/) から Aspose.Cells Cloud Perl モジュールをインストールします。
3. 以下の変換コードを参考にして、アプリケーションを追加または変更します。

## Perl を使用した Excel ファイルの変換

```perl
use strict;
use warnings;
use File::Slurp;
use MIME::Base64;
use AsposeCellsCloud::ApiClient;
use AsposeCellsCloud::CellsApi;
use AsposeCellsCloud::Configuration;
use AsposeCellsCloud::Request::ConvertSpreadsheetRequest;

my $config = AsposeCellsCloud::Configuration->new( client_id => $ENV{'CellsCloudClientId'}, client_secret => $ENV{'CellsCloudClientSecret'});
my $instance = AsposeCellsCloud::CellsApi->new(AsposeCellsCloud::ApiClient->new( $config));

my $request = AsposeCellsCloud::Request::ConvertSpreadsheetRequest->new();
$request->{spreadsheet} =  'EmployeeSalesSummary.xlsx';
$request->{format} = 'pdf';
my $response = $instance->put_convert_workbook(request=> $request);
open (my $fh, '>', 'EmployeeSalesSummary.pdf') or die "Cannot open EmployeeSalesSummary.pdf $!";
binmode($fh);
print $fh $response;
close($fh);
```

## サポートされているファイル形式

| **形式**                                                          | **説明**                                                                                                                     | **読み込み** | **保存** |
| :---------------------------------------------------------------- | :--------------------------------------------------------------------------------------------------------------------------- | :----------- | :------- |
| [XLS](https://docs.fileformat.com/spreadsheet/xls/)               | Excel 95/5.0 - 2003 ワークブック。                                                                                            | &radic;      | &radic;  |
| [XLSX](https://docs.fileformat.com/spreadsheet/xlsx/)             | Office Open XML SpreadsheetML ワークブックまたはテンプレートファイル（マクロの有無は問いません）。                                | &radic;      | &radic;  |
| [XLSB](https://docs.fileformat.com/spreadsheet/xlsb/)             | Excel バイナリワークブック。                                                                                                  | &radic;      | &radic;  |
| [XLSM](https://docs.fileformat.com/spreadsheet/xlsm/)             | Excel マクロ有効ワークブック。                                                                                                | &radic;      | &radic;  |
| [XLT](https://docs.fileformat.com/spreadsheet/xlt/)               | Excel 97 - Excel 2003 テンプレート。                                                                                          | &radic;      | &radic;  |
| [XLTX](https://docs.fileformat.com/spreadsheet/xltx/)             | Excel テンプレート。                                                                                                          | &radic;      | &radic;  |
| [XLTM](https://docs.fileformat.com/spreadsheet/xltm/)             | Excel マクロ有効テンプレート。                                                                                                | &radic;      | &radic;  |
| [XLAM](https://docs.fileformat.com/spreadsheet/xlam/)             | Excel に新しい機能を追加するために使用される Excel マクロ有効アドインファイル。                                                  |              | &radic;  |
| [CSV](https://docs.fileformat.com/spreadsheet/csv/)               | CSV（カンマ区切り値）ファイル。                                                                                                | &radic;      | &radic;  |
| [TSV](https://docs.fileformat.com/spreadsheet/tsv/)               | TSV（タブ区切り値）ファイル。                                                                                                 | &radic;      | &radic;  |
| [TXT](https://docs.fileformat.com/word-processing/txt/)           | 区切り文字付きプレーンテキストファイル。                                                                                      | &radic;      | &radic;  |
| [HTML](https://docs.fileformat.com/web/html/)                     | HTML 形式。                                                                                                                   | &radic;      | &radic;  |
| [MHTML](https://docs.fileformat.com/web/mhtml/)                   | MHTML ファイル。                                                                                                              | &radic;      | &radic;  |
| [ODS](https://docs.fileformat.com/spreadsheet/ods/)               | ODS（OpenDocument スプレッドシート）。                                                                                         | &radic;      | &radic;  |
| [Numbers](https://docs.fileformat.com/spreadsheet/numbers/)       | Apple の iWork オフィススイートの一部である Apple の "Numbers" アプリケーションで作成されたドキュメントで、Mac OS X および iOS で動作します。 | &radic;      |          |
| [JSON](https://docs.fileformat.com/web/json/)                     | JavaScript Object Notation。                                                                                                  | &radic;      | &radic;  |
| [DIF](https://docs.fileformat.com/spreadsheet/dif/)               | Data Interchange Format。                                                                                                     |              | &radic;  |
| [PDF](https://docs.fileformat.com/pdf/)                           | Adobe Portable Document Format。                                                                                              |              | &radic;  |
| [XPS](https://docs.fileformat.com/page-description-language/xps/) | XML Paper Specification Format。                                                                                              |              | &radic;  |
| [SVG](https://docs.fileformat.com/page-description-language/svg/) | Scalable Vector Graphics Format。                                                                                             |              | &radic;  |
| [TIFF](https://docs.fileformat.com/image/tiff/)                   | Tagged Image File Format。                                                                                                    |              | &radic;  |
| [PNG](https://docs.fileformat.com/image/png/)                     | Portable Network Graphics Format。                                                                                            |              | &radic;  |
| [BMP](https://docs.fileformat.com/image/bmp/)                     | Bitmap Image Format.                                                                                                          |              | &radic;  |
| [EMF](https://docs.fileformat.com/image/emf/)                     | Enhanced Metafile Format。                                                                                                    |              | &radic;  |
| [JPEG](https://docs.fileformat.com/image/jpeg/)                   | JPEG は非可逆圧縮を使用する画像形式です。                                                                                    |              | &radic;  |
| [GIF](https://docs.fileformat.com/image/gif/)                     | Graphical Interchange Format。                                                                                                |              | &radic;  |
| [MARKDOWN](https://docs.fileformat.com/word-processing/md/)       | Markdown ドキュメントを表します。                                                                                              |              | &radic;  |
| [SXC](https://docs.fileformat.com/spreadsheet/sxc/)               | OpenOffice および StarOffice で使用される XML ベースの形式。                                                                   | &radic;      | &radic;  |
| [FODS](https://docs.fileformat.com/spreadsheet/fods/)             | フラット XML として保存された Open Document 形式。                                                                             | &radic;      | &radic;  |
| [DOCX](https://docs.fileformat.com/word-processing/docx/)         | XML とバイナリファイルを組み合わせた Microsoft Word ドキュメントのよく知られた形式。                                            |              | &radic;  |
| [PPTX](https://docs.fileformat.com/presentation/pptx/)            | Microsoft PowerPoint Open XML プレゼンテーションファイル形式に基づいています。                                                   |              | &radic;  |
| [OTS](https://docs.fileformat.com/spreadsheet/ots/)               | OTS（OpenDocument スプレッドシート）。                                                                                         | &radic;      | &radic;  |
| [XML](https://docs.fileformat.com/web/xml/)                       | XML ファイル。                                                                                                                | &radic;      | &radic;  |
| [HTM](https://docs.fileformat.com/web/htm/)                       | HTM ファイル。                                                                                                                | &radic;      | &radic;  |
| [TIF](https://docs.fileformat.com/image/tiff/)                    | Tagged Image File Format。                                                                                                    |              | &radic;  |
| [WMF](https://docs.fileformat.com/image/wmf/)                     | Windows Metafile Format。                                                                                                     |              | &radic;  |
| [PCL](https://docs.fileformat.com/page-description-language/pcl/) | Printer Command Language Format。                                                                                             |              | &radic;  |
| [AZW3](https://docs.fileformat.com/ebook/azw3/)                   | AZW3/KF8 ファイル形式。                                                                                                       |              | &radic;  |
| [EPUB](https://docs.fileformat.com/ebook/epub/)                   | EPUB ファイル形式。                                                                                                           |              | &radic;  |
| [DBF](https://docs.fileformat.com/database/dbf/)                  | データベースファイル形式。                                                                                                    |              | &radic;  |
| [XHTML](https://docs.fileformat.com/web/xhtml/)                   | XHTML ファイル形式。                                                                                                          |              | &radic;  |

## アーキテクチャ

![architecture.png](architecture.png)

## [開発者リファレンス](docs/DeveloperGuide.md#overview)

### クラウドでの Excel およびその他のスプレッドシートファイルの操作

- **ファイル操作**: クラウドに保存された Excel ファイルのアップロード、ダウンロード、削除、管理が可能です。
- **ファイル書式設定**: 特定の要件に合わせて、Excel ファイルのセル、フォント、色、配置モードの書式設定をサポートします。
- **データ処理**: セルデータの読み取り、書き込み、変更、数式計算の実行、データの書式設定など、強力なデータ処理機能を提供します。
- **数式計算**: 組み込みの数式エンジンが Excel の複雑な数式計算を処理し、正確な結果を返します。
- **グラフ操作**: データ分析や可視化のニーズに応じて、Excel ファイルのグラフを作成、編集、削除できます。
- **テーブル処理**: 作成、編集、書式設定、変換など、さまざまな形式操作に対する堅牢な処理機能を提供します。
- **データ検証**: セルのデータ型、範囲、一意性を設定するデータ検証機能を備え、データの正確性と整合性を確保します。
- **バッチ処理**: バッチ形式変換、データ抽出、スタイル適用など、複数の Excel ドキュメントのバッチ処理をサポートします。
- **データのインポート/エクスポート**: さまざまなソースからスプレッドシートへのデータインポート、およびスプレッドシートデータの他の形式へのエクスポートを容易にします。
- **セキュリティ管理**: データ暗号化、アクセス制御、権限管理などのセキュリティ機能を提供し、スプレッドシートデータのセキュリティと整合性を保護します。

## バージョン 26.8 の機能と拡張

| **概要** | **カテゴリ** |
| :--- | :--- |
| Enchent smart template feature. | Improvement |
| A new AI data analysis API has been added. | New Feature |

## 利用可能な SDK

Aspose.Cells Cloud SDK は複数の一般的なプログラミング言語で利用可能で、開発者はさまざまな開発環境でスプレッドシート処理機能を統合できます。

[![Go](https://img.shields.io/badge/Go-00ADD8.svg?style=for-the-badge&logo=go&logoColor=white)](https://github.com/aspose-cells-cloud/aspose-cells-cloud-go) [![Go](https://img.shields.io/badge/Go-Install%20go%20get%20package--asposecellscloud-%2300ADD8?logo=go&style=for-the-badge)](https://pkg.go.dev/github.com/aspose-cells-cloud/aspose-cells-cloud-go/v25)

[![Java](https://img.shields.io/badge/Java-red?logo=openjdk&style=for-the-badge&logoColor=white)](https://github.com/aspose-cells-cloud/aspose-cells-cloud-java) [![Java](https://img.shields.io/badge/Maven-Aspose.Cells%20Cloud.pom.xml-red?logo=apache-maven&style=for-the-badge)](https://github.com/aspose-cells-cloud/aspose-cells-cloud-java/blob/master/Aspose.Cells.Cloud.pom.xml)

[![.NET](https://img.shields.io/badge/.NET-%23512BD4?style=for-the-badge&logo=dotnet&logoColor=white)](https://github.com/aspose-cells-cloud/aspose-cells-cloud-dotnet) [![.NET](https://img.shields.io/badge/NuGet-Install%20Aspose.Cells--Cloud-%23512BD4?logo=nuget&style=for-the-badge)](https://www.nuget.org/packages/Aspose.cells-Cloud/#readme-body-tab)

[![Node.js](https://img.shields.io/badge/Node.js-43853D.svg?style=for-the-badge&logo=node.js&logoColor=white)](https://github.com/aspose-cells-cloud/aspose-cells-cloud-node) [![Node.js](https://img.shields.io/badge/npm-install%20asposecellscloud-orange?logo=npm&style=for-the-badge)](https://www.npmjs.com/package/asposecellscloud)

[![Perl](https://img.shields.io/badge/Perl-39457E.svg?style=for-the-badge&logo=perl&logoColor=white)](https://github.com/aspose-cells-cloud/aspose-cells-cloud-perl) [![Perl](https://img.shields.io/badge/CPAN-Install%20AsposeCellsCloud--CellsApi-blue?logo=perl&style=for-the-badge)](https://metacpan.org/dist/AsposeCellsCloud-CellsApi)

[![PHP](https://img.shields.io/badge/PHP-777BB4.svg?style=for-the-badge&logo=php&logoColor=white)](https://github.com/aspose-cells-cloud/aspose-cells-cloud-php) [![PHP](https://img.shields.io/badge/Composer-require%20aspose/cells--sdk--php-8892BF?logo=php&style=for-the-badge)](https://packagist.org/packages/aspose/cells-sdk-php)

[![Python](https://img.shields.io/badge/Python-14354C.svg?style=for-the-badge&logo=python&logoColor=white)](https://github.com/aspose-cells-cloud/aspose-cells-cloud-python) [![Python](https://img.shields.io/badge/pip-install%20asposecellscloud-blue?logo=pypi&style=for-the-badge)](https://pypi.org/project/asposecellscloud/)

[![Ruby](https://img.shields.io/badge/Ruby-CC342D.svg?style=for-the-badge&logo=ruby&logoColor=white)](https://github.com/aspose-cells-cloud/aspose-cells-cloud-ruby) [![Ruby](https://img.shields.io/badge/Gem-install%20aspose__cells__cloud-red?logo=ruby&style=for-the-badge)](https://rubygems.org/gems/aspose_cells_cloud)

## [リリース履歴](CHANGELOG.md)

---

<h1 id="deutsch">Deutsch</h1>

[Aspose.Cells Cloud SDK for Perl](https://products.aspose.cloud/cells/perl) ist eine cloud-native REST-API, die Perl-Entwicklern ermöglicht, Tabellenkalkulationsdateien — einschließlich **Excel** (**XLS**, **XLSX**, **XLSB**, **XLSM**), **OpenDocument Spreadsheet** (**ODS**), **CSV**, **TSV**, **JSON**, **HTML**, **PDF** und **mehr — ohne Installation von Microsoft Excel oder Office zu** **erstellen**, **zu lesen**, **zu bearbeiten**, **zu konvertieren** und **zu reparieren**.

Aufbauend auf der **Aspose.Cells Cloud Web API** unterstützt dieses MIT-lizenzierte SDK erweiterte Tabellenkalkulationsoperationen wie:

- Zellenformatierung, Formeln und Datenvalidierung
- Pivot-Tabellen, Diagramme, Hyperlinks und Kommentare
- Bedingte Formatierung und Smart Marker
- Arbeitsblattzusammenführung, -aufteilung und -schutz
- Stapelverarbeitung und Hintergrundentfernung

Es integriert sich nahtlos in **AWS**, **Microsoft Azure** und **Google Cloud** und gewährleistet **hohe Verfügbarkeit**, **Skalierbarkeit** und **Datenintegrität**. Ideal für serverlose Anwendungen, Microservices und Cloud-Automatisierungsworkflows.

## Kurzanleitung

Um mit Aspose.Cells Cloud zu beginnen, führen Sie folgende Schritte aus:

1. Registrieren Sie sich bei [Aspose for Cloud](https://dashboard.aspose.cloud/#/apps) für ein Konto, um Ihre Anwendungsdaten zu erhalten.
2. Installieren Sie das Aspose.Cells Cloud Perl-Modul aus der [CPAN-Distribution](https://www.cpan.org/).
3. Verwenden Sie den unten angegebenen Konvertierungscode als Referenz, um Ihre Anwendung hinzuzufügen oder zu ändern.

## Konvertieren einer Excel-Datei mit Perl

```perl
use strict;
use warnings;
use File::Slurp;
use MIME::Base64;
use AsposeCellsCloud::ApiClient;
use AsposeCellsCloud::CellsApi;
use AsposeCellsCloud::Configuration;
use AsposeCellsCloud::Request::ConvertSpreadsheetRequest;

my $config = AsposeCellsCloud::Configuration->new( client_id => $ENV{'CellsCloudClientId'}, client_secret => $ENV{'CellsCloudClientSecret'});
my $instance = AsposeCellsCloud::CellsApi->new(AsposeCellsCloud::ApiClient->new( $config));

my $request = AsposeCellsCloud::Request::ConvertSpreadsheetRequest->new();
$request->{spreadsheet} =  'EmployeeSalesSummary.xlsx';
$request->{format} = 'pdf';
my $response = $instance->put_convert_workbook(request=> $request);
open (my $fh, '>', 'EmployeeSalesSummary.pdf') or die "Cannot open EmployeeSalesSummary.pdf $!";
binmode($fh);
print $fh $response;
close($fh);
```

## Unterstützte Dateiformate

| **Format**                                                        | **Beschreibung**                                                                                                                          | **Laden** | **Speichern** |
| :---------------------------------------------------------------- | :---------------------------------------------------------------------------------------------------------------------------------------- | :-------- | :------------ |
| [XLS](https://docs.fileformat.com/spreadsheet/xls/)               | Excel 95/5.0 - 2003-Arbeitsmappe.                                                                                                         | &radic;   | &radic;       |
| [XLSX](https://docs.fileformat.com/spreadsheet/xlsx/)             | Office Open XML SpreadsheetML-Arbeitsmappe oder Vorlagendatei, mit oder ohne Makros.                                                      | &radic;   | &radic;       |
| [XLSB](https://docs.fileformat.com/spreadsheet/xlsb/)             | Excel-Binärarbeitsmappe.                                                                                                                  | &radic;   | &radic;       |
| [XLSM](https://docs.fileformat.com/spreadsheet/xlsm/)             | Excel-Arbeitsmappe mit Makros.                                                                                                            | &radic;   | &radic;       |
| [XLT](https://docs.fileformat.com/spreadsheet/xlt/)               | Excel 97 - Excel 2003-Vorlage.                                                                                                            | &radic;   | &radic;       |
| [XLTX](https://docs.fileformat.com/spreadsheet/xltx/)             | Excel-Vorlage.                                                                                                                            | &radic;   | &radic;       |
| [XLTM](https://docs.fileformat.com/spreadsheet/xltm/)             | Excel-Vorlage mit Makros.                                                                                                                 | &radic;   | &radic;       |
| [XLAM](https://docs.fileformat.com/spreadsheet/xlam/)             | Excel-Add-In-Datei mit Makros, die zum Hinzufügen neuer Funktionen zu Excel verwendet wird.                                              |           | &radic;       |
| [CSV](https://docs.fileformat.com/spreadsheet/csv/)               | CSV (Comma Separated Value)-Datei.                                                                                                        | &radic;   | &radic;       |
| [TSV](https://docs.fileformat.com/spreadsheet/tsv/)               | TSV (Tab-separated values)-Datei.                                                                                                         | &radic;   | &radic;       |
| [TXT](https://docs.fileformat.com/word-processing/txt/)           | Durch Trennzeichen getrennte reine Textdatei.                                                                                             | &radic;   | &radic;       |
| [HTML](https://docs.fileformat.com/web/html/)                     | HTML-Format.                                                                                                                              | &radic;   | &radic;       |
| [MHTML](https://docs.fileformat.com/web/mhtml/)                   | MHTML-Datei.                                                                                                                              | &radic;   | &radic;       |
| [ODS](https://docs.fileformat.com/spreadsheet/ods/)               | ODS (OpenDocument Spreadsheet).                                                                                                           | &radic;   | &radic;       |
| [Numbers](https://docs.fileformat.com/spreadsheet/numbers/)       | Dokumente, die mit Apples Anwendung "Numbers" erstellt wurden, die Teil von Apples iWork-Office-Suite ist und unter Mac OS X und iOS läuft. | &radic;   |               |
| [JSON](https://docs.fileformat.com/web/json/)                     | JavaScript Object Notation.                                                                                                               | &radic;   | &radic;       |
| [DIF](https://docs.fileformat.com/spreadsheet/dif/)               | Data Interchange Format.                                                                                                                  |           | &radic;       |
| [PDF](https://docs.fileformat.com/pdf/)                           | Adobe Portable Document Format.                                                                                                            |           | &radic;       |
| [XPS](https://docs.fileformat.com/page-description-language/xps/) | XML Paper Specification Format.                                                                                                            |           | &radic;       |
| [SVG](https://docs.fileformat.com/page-description-language/svg/) | Scalable Vector Graphics Format.                                                                                                           |           | &radic;       |
| [TIFF](https://docs.fileformat.com/image/tiff/)                   | Tagged Image File Format.                                                                                                                 |           | &radic;       |
| [PNG](https://docs.fileformat.com/image/png/)                     | Portable Network Graphics Format.                                                                                                          |           | &radic;       |
| [BMP](https://docs.fileformat.com/image/bmp/)                     | Bitmap Image Format.                                                                                                                       |           | &radic;       |
| [EMF](https://docs.fileformat.com/image/emf/)                     | Enhanced Metafile Format.                                                                                                                  |           | &radic;       |
| [JPEG](https://docs.fileformat.com/image/jpeg/)                   | JPEG ist ein Bildformat, das verlustbehaftete Komprimierung verwendet.                                                                    |           | &radic;       |
| [GIF](https://docs.fileformat.com/image/gif/)                     | Graphical Interchange Format.                                                                                                             |           | &radic;       |
| [MARKDOWN](https://docs.fileformat.com/word-processing/md/)       | Stellt ein Markdown-Dokument dar.                                                                                                         |           | &radic;       |
| [SXC](https://docs.fileformat.com/spreadsheet/sxc/)               | Ein XML-basiertes Format, das von OpenOffice und StarOffice verwendet wird.                                                               | &radic;   | &radic;       |
| [FODS](https://docs.fileformat.com/spreadsheet/fods/)             | Ein als flaches XML gespeichertes Open Document-Format.                                                                                   | &radic;   | &radic;       |
| [DOCX](https://docs.fileformat.com/word-processing/docx/)         | Ein bekanntes Format für Microsoft Word-Dokumente, das XML- und Binärdateien kombiniert.                                                   |           | &radic;       |
| [PPTX](https://docs.fileformat.com/presentation/pptx/)            | Das PPTX-Format basiert auf dem Microsoft PowerPoint Open XML-Präsentationsdateiformat.                                                    |           | &radic;       |
| [OTS](https://docs.fileformat.com/spreadsheet/ots/)               | OTS (OpenDocument Spreadsheet).                                                                                                           | &radic;   | &radic;       |
| [XML](https://docs.fileformat.com/web/xml/)                       | XML-Datei.                                                                                                                                | &radic;   | &radic;       |
| [HTM](https://docs.fileformat.com/web/htm/)                       | HTM-Datei.                                                                                                                                | &radic;   | &radic;       |
| [TIF](https://docs.fileformat.com/image/tiff/)                    | Tagged Image File Format.                                                                                                                 |           | &radic;       |
| [WMF](https://docs.fileformat.com/image/wmf/)                     | Windows Metafile Format.                                                                                                                  |           | &radic;       |
| [PCL](https://docs.fileformat.com/page-description-language/pcl/) | Printer Command Language Format.                                                                                                          |           | &radic;       |
| [AZW3](https://docs.fileformat.com/ebook/azw3/)                   | AZW3/KF8-Dateiformat.                                                                                                                     |           | &radic;       |
| [EPUB](https://docs.fileformat.com/ebook/epub/)                   | EPUB-Dateiformat.                                                                                                                         |           | &radic;       |
| [DBF](https://docs.fileformat.com/database/dbf/)                  | Datenbank-Dateiformat.                                                                                                                    |           | &radic;       |
| [XHTML](https://docs.fileformat.com/web/xhtml/)                   | XHTML-Dateiformat.                                                                                                                        |           | &radic;       |

## Architektur

![architecture.png](architecture.png)

## [Entwicklerreferenz](docs/DeveloperGuide.md#overview)

### Bearbeiten von Excel und anderen Tabellenkalkulationsdateien in der Cloud

- **Dateimanipulation**: Benutzer können in der Cloud gespeicherte Excel-Dateien hochladen, herunterladen, löschen und verwalten.
- **Dateiformatierung**: Unterstützt die Formatierung von Zellen, Schriftarten, Farben und Ausrichtungsmodi in Excel-Dateien entsprechend spezifischer Anforderungen.
- **Datenverarbeitung**: Leistungsstarke Funktionen zur Datenverarbeitung, einschließlich Lesen, Schreiben, Ändern von Zelldaten, Durchführen von Formelberechnungen und Formatieren von Daten.
- **Formelberechnung**: Die integrierte Formel-Engine verarbeitet komplexe Formelberechnungen in Excel und liefert genaue Ergebnisse.
- **Diagrammbearbeitung**: Benutzer können Diagramme aus Excel-Dateien für Datenanalyse- und Visualisierungsanforderungen erstellen, bearbeiten und löschen.
- **Tabellenverarbeitung**: Bietet robuste Verarbeitungsfunktionen für verschiedene Formatoperationen wie Erstellung, Bearbeitung, Formatierung und Konvertierung.
- **Datenvalidierung**: Enthält Datenvalidierungsfunktionen zum Festlegen von Zelldatentyp, Bereich und Eindeutigkeit, um Datengenauigkeit und -integrität sicherzustellen.
- **Stapelverarbeitung**: Unterstützt die Stapelverarbeitung mehrerer Excel-Dokumente, wie z. B. Stapelformatkonvertierung, Datenextraktion und Stilanwendung.
- **Datenimport/-export**: Ermöglicht den Import von Daten aus verschiedenen Quellen in Tabellen und den Export von Tabellendaten in andere Formate.
- **Sicherheitsmanagement**: Bietet eine Reihe von Sicherheitsfunktionen wie Datenverschlüsselung, Zugriffskontrolle und Berechtigungsverwaltung zum Schutz der Sicherheit und Integrität von Tabellendaten.

## Funktionen & Erweiterungen in Version 26.8

| **Zusammenfassung**                                             | **Kategorie**  |
| :-------------------------------------------------------------- | :------------- |
| Enchent smart template feature.                                 | Improvement |
| A new AI data analysis API has been added.                      | New Feature |

## Verfügbare SDKs

Das Aspose.Cells Cloud SDK ist in mehreren gängigen Programmiersprachen verfügbar und ermöglicht Entwicklern die Integration von Tabellenverarbeitungsfunktionen in verschiedene Entwicklungsumgebungen.

[![Go](https://img.shields.io/badge/Go-00ADD8.svg?style=for-the-badge&logo=go&logoColor=white)](https://github.com/aspose-cells-cloud/aspose-cells-cloud-go) [![Go](https://img.shields.io/badge/Go-Install%20go%20get%20package--asposecellscloud-%2300ADD8?logo=go&style=for-the-badge)](https://pkg.go.dev/github.com/aspose-cells-cloud/aspose-cells-cloud-go/v25)

[![Java](https://img.shields.io/badge/Java-red?logo=openjdk&style=for-the-badge&logoColor=white)](https://github.com/aspose-cells-cloud/aspose-cells-cloud-java) [![Java](https://img.shields.io/badge/Maven-Aspose.Cells%20Cloud.pom.xml-red?logo=apache-maven&style=for-the-badge)](https://github.com/aspose-cells-cloud/aspose-cells-cloud-java/blob/master/Aspose.Cells.Cloud.pom.xml)

[![.NET](https://img.shields.io/badge/.NET-%23512BD4?style=for-the-badge&logo=dotnet&logoColor=white)](https://github.com/aspose-cells-cloud/aspose-cells-cloud-dotnet) [![.NET](https://img.shields.io/badge/NuGet-Install%20Aspose.Cells--Cloud-%23512BD4?logo=nuget&style=for-the-badge)](https://www.nuget.org/packages/Aspose.cells-Cloud/#readme-body-tab)

[![Node.js](https://img.shields.io/badge/Node.js-43853D.svg?style=for-the-badge&logo=node.js&logoColor=white)](https://github.com/aspose-cells-cloud/aspose-cells-cloud-node) [![Node.js](https://img.shields.io/badge/npm-install%20asposecellscloud-orange?logo=npm&style=for-the-badge)](https://www.npmjs.com/package/asposecellscloud)

[![Perl](https://img.shields.io/badge/Perl-39457E.svg?style=for-the-badge&logo=perl&logoColor=white)](https://github.com/aspose-cells-cloud/aspose-cells-cloud-perl) [![Perl](https://img.shields.io/badge/CPAN-Install%20AsposeCellsCloud--CellsApi-blue?logo=perl&style=for-the-badge)](https://metacpan.org/dist/AsposeCellsCloud-CellsApi)

[![PHP](https://img.shields.io/badge/PHP-777BB4.svg?style=for-the-badge&logo=php&logoColor=white)](https://github.com/aspose-cells-cloud/aspose-cells-cloud-php) [![PHP](https://img.shields.io/badge/Composer-require%20aspose/cells--sdk--php-8892BF?logo=php&style=for-the-badge)](https://packagist.org/packages/aspose/cells-sdk-php)

[![Python](https://img.shields.io/badge/Python-14354C.svg?style=for-the-badge&logo=python&logoColor=white)](https://github.com/aspose-cells-cloud/aspose-cells-cloud-python) [![Python](https://img.shields.io/badge/pip-install%20asposecellscloud-blue?logo=pypi&style=for-the-badge)](https://pypi.org/project/asposecellscloud/)

[![Ruby](https://img.shields.io/badge/Ruby-CC342D.svg?style=for-the-badge&logo=ruby&logoColor=white)](https://github.com/aspose-cells-cloud/aspose-cells-cloud-ruby) [![Ruby](https://img.shields.io/badge/Gem-install%20aspose__cells__cloud-red?logo=ruby&style=for-the-badge)](https://rubygems.org/gems/aspose_cells_cloud)

## [Versionsverlauf](CHANGELOG.md)