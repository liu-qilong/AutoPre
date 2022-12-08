# AutoPre

Thank you for using AutoPre, a ppt automatic filler based on [Dynalist note](https://dynalist.io), [MS Word document](https://www.microsoft.com/en-ww/microsoft-365/word) and so on, targeting on freeing ourselves from the exhausting ppt formating. Thus, we can focus on the content itself and contribute more valuable outputs.
 
## Dependencies

These scripts are run in Python environment. We assume that you have installed it successfully. The whole project is also based on Python libraries, [python-pptx](https://python-pptx.readthedocs.io/en/latest/index.html) and [python-docx](https://python-docx.readthedocs.io/en/latest/), for manipulating `.pptx` and `.docx` files. Hence, the dependencies of AutoPre are shown below:
 
```
Python 2.6, 2.7 or 3.x
python-pptx
python-docx
lxml >= 2.3.2
Pillow
XlsxWriter
```

## Usage

### Content Preparation

Concentrating on the content itself and the scripts automatically filling the content into the PowerPoint template, that's the whole idea. Thus, we need a carrier for the content. Dynalist and MS Word both serve well for this purpose, therefore they are selected. Put them into the `AutoPre/documents/` folder and the scripts can detect and load them automatically.
 
#### Dynalist Documents Preparation

Dynalist is an outline note software. In AutoPre, the first level nodes of it will become first level section slides. And if their notes are not empty, makes text slides following. Then the second level nodes will become second level section slides or text slides, depending on your choice, which we will mention in the next section. Other level nodes will become text slides.

All the text slides can be mounted with image illustration areas, depending on your choice, which we will mention in the next section.

Please note that the Dynalist documents should be [exported](https://help.dynalist.io/category/36-import-export) as `.opml` files and put into the `AutoPre/documents` folder.
  
![Sample of .opml file](https://github.com/TOB-KNPOB/AutoPre/blob/main/gallery/sample_of_opml.png)
  
#### MS Word Documents Preparation

In MS Word we can edit the style of a word conveniently. [Specifically, on the Home tab left-click any style in the Styles gallery.](https://support.microsoft.com/en-us/office/customize-or-create-new-styles-d38d6e47-f6fc-48eb-a607-1eb120dec563#:~:text=On%20the%20Home%20tab%2C%20right,or%20to%20all%20future%20documents.) AutoPre uses the styles to divide the whole article into different hierarchies.

In AutoPre, the paragraphs whose styles are "Heading 1" will become first level section slides. And if there are paragraphs whose styles are "Normal" following them, they will become text slides. Then the paragraphs whose styles are "Heading 2" will become second level section slides or text slides, depending on your choice, which we will mention in the next section. The paragraphs whose styles are "Heading 3" will become the title of the text slides. And the paragraphs whose styles are "Normal" will become the contents following the titles.

All the text slides can be mounted with image illustration areas, depending on your choice, which we will mention in the next section.

Please note that the MS Word documents should be saved as `.docx` files and put into the `AutoPre/documents` folder.
 
![Sample of .docx file](https://github.com/TOB-KNPOB/AutoPre/blob/main/gallery/sample_of_docx.png)

### Run the Script

Run the AutoPre.py script in the root directory and the guidance will be shown in the console area. It will guide you through the documents loading, parameters setting and PowerPoint (.pptx) files outputting.

The guidance will tell you how to define the title of this presentation, date and author, offer you with choices of templates, let you determine whether the second level sections slides will be used, and whether the image illustration will be used.
 Eventually, AutoPre will generate a PowerPoint `.pptx` documents as required.
 Here is an example of the console command lines and its outputs. 
 
```
AutoPre:
  Thank you for using AutoPre, a PPT automatic filler based on Dynalist note, Word document and so on, targeting on freeing ourselves from the exhausting ppt formating. Thus, we can focus on the content itself and contribute more valuable outputs.

  感谢您使用AutoPre，一款基于Dynalist笔记、Word文档等的ppt自动填充软件，旨在让我们从繁重的ppt制作中解脱出来. 这样，我们就可以专注于内容本身，贡献更多有价值的产出.

AutoPre:
  DocxExample.docx has been loaded.
  OpmlExample.opml has been loaded.

  2 files has been loaded.
  2份文件被加载.

AutoPre:
  Setting the DocxExample.docx's output parameters.
  设置 DocxExample.docx 的输出参数.

  Please enter the name of this presentation.
  请输入该演示文档的标题.
  Docx Example

  Please enter the date statement of presentation. eg.Date  7/11/1917
  请输入演示文档的日期说明. 示例:日期  1917/11/7
  Date  30/7/2021

  Please enter the inscribe includes the author's name. eg.Reporter Kirov
  请输入包含作者姓名的落款. 示例:汇报  李明
  Report  Logic Flow

  Please enter the template name. Here are the choices:
  请输入模板名称，以下为选择范围:
  SWU  SZU  PolyU  Marxist  
  PolyU

  Would you like to use second level sections pages?
  Please enter 'Y' representing yes or 'N' representing no.
  您希望使用二级章节分界页面吗？
  请输入 Y 代表是，输入 N 代表否.
  Y

  Would you like to add image illustration to each main body pages?
  Please enter 'Y' representing yes or 'N' representing no.
  您希望给正文页面添加插图区吗？
  请输入 Y 代表是，输入 N 代表否.
  Y

  Presentation document is outputted to outputs/Docx Example.pptx
  演示文档输出到 outputs/Docx Example.pptx

AutoPre:
  Setting the OpmlExample.opml's output parameters.
  设置 OpmlExample.opml 的输出参数.

  Please enter the name of this presentation.
  请输入该演示文档的标题.
  Opml Example

  Please enter the date statement of presentation. eg.Date  7/11/1917
  请输入演示文档的日期说明. 示例:日期  1917/11/7
  Date  30/7/2021

  Please enter the inscribe includes the author's name. eg.Reporter Kirov
  请输入包含作者姓名的落款. 示例:汇报  李明
  Report  Logic Flow

  Please enter the template name. Here are the choices:
  请输入模板名称，以下为选择范围:
  SWU  SZU  PolyU  Marxist  
  SZU

  Would you like to use second level sections pages?
  Please enter 'Y' representing yes or 'N' representing no.
  您希望使用二级章节分界页面吗？
  请输入 Y 代表是，输入 N 代表否.
  N

  Would you like to add image illustration to each main body pages?
  Please enter 'Y' representing yes or 'N' representing no.
  您希望给正文页面添加插图区吗？
  请输入 Y 代表是，输入 N 代表否.
  N

  Presentation document is outputted to outputs/Opml Example.pptx
  演示文档输出到 outputs/Opml Example.pptx
```

 ![Sample of cover slide](https://github.com/TOB-KNPOB/AutoPre/blob/main/gallery/sample_of_cover_slide_PolyU.png)

 ![Sample of cover slide](https://github.com/TOB-KNPOB/AutoPre/blob/main/gallery/sample_of_cover_slide_SZU.png)

 ![Sample of first level section slide](https://github.com/TOB-KNPOB/AutoPre/blob/main/gallery/sample_of_section_1st_level.png)

 ![Sample of second level section slide](https://github.com/TOB-KNPOB/AutoPre/blob/main/gallery/sample_of_section_2nd_level.png)

 ![Sample of pure text slide](https://github.com/TOB-KNPOB/AutoPre/blob/main/gallery/sample_of_pure_text_slide.png)

 ![Sample of text with image slide](https://github.com/TOB-KNPOB/AutoPre/blob/main/gallery/sample_of_text_with_image_slide.png)

 ![Sample of back cover slide](https://github.com/TOB-KNPOB/AutoPre/blob/main/gallery/sample_of_back_cover_slide.png)

 
### Checking and Composing

Then you can open the generated documents and edit them like any usual PowerPoint document. You can change the layout of the elements, fonts of the text, combine the slides that are too sparse, split the slides that are too dense, and whatever you want.
 
## About the Templates

We offer a few formal templates for you. You can edit it and change it as you like. But there are three things to be noted:

- You need to enter [the Slide Master](https://support.microsoft.com/en-us/office/edit-and-re-apply-a-slide-layout-6f4338f8-555f-49cf-9835-6209be3c7b48) view to edit the template.
- Please don't delete any placeholders in it. But you can change the fonts, color, typeface as you like.
- Please add the theme name and theme color into the `THEME_AND_COLORS` dictionary at the beginning of `AutoPre\instruments\strategy.py` file.
eg. Adding a template named `MyTemplate.pptx` with a theme color in [RGB format](https://en.wikipedia.org/wiki/RGB_color_model). Then the definition of `THEME_AND_COLORS` should be modified to:

```
THEME_AND_COLORS = {
  'SWU': [33,39,113],
  'SZU': [145,25,65],
  'PolyU': [161,35,56],
  'Marxist': [160,1,2],
  'MyTemplate': [0,0,0]
  }
```
