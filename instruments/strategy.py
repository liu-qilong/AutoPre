from . import load
from . import slidemaker
from pptx import Presentation

# The theme colors of the templates

THEME_AND_COLORS = {
    'SWU': [33,39,113],
    'SZU': [145,25,65],
    'PolyU': [161,35,56],
    'Marxist': [160,1,2]
    }


def standard_strategy(node, prs_info, prs, level_mode, img_mode):
    """The standard strategy for ppt formated filling.
    
    Noted that the function outputs the information from a nodes' tree.
    Therefore, this function is called traversally.
    
    Args:
        node: The OutlinNode object that this slide contains.
        prs_info: A dictory that contains the basic information of the presentation for content filling.
            title: A string contains the title of the presentation.
            date: A string contains the data of the presentation, usually formatted in "Date [day/month/year]".
            part_num: A string that indicates the serial number of the current part, usually formatted in "Part [num]".
            part_title: A string contains the title of the current part.
            color: A list that contains the main theme color in RGB format.
            author: A string contains the name of the author, usually formatted in "Report [name]".
        prs: The presentation object it works on.
        level_mode: A boolean indicates if we use second level sections or not.
        img_mode: A boolean indicates if we use image in text pages or not.
        
    """
    
    # If the node's layer is 0, makes a cover slide.
    
    if node.layer == 0:
        
        slidemaker.CoverSlideMaker(node, prs_info, prs)
        
    # If the node's layer is 1, makes a first level section slide
    # And if the node's note is not empty, makes a pure text slide or text with image slide following it.
        
    elif node.layer == 1:
        
        prs_info['part_num'] +=  1
        prs_info['part_title'] = node.title
        
        slidemaker.SectionSlideMaker(node, prs_info, prs, 1)
        
        if node.note != '':
            if img_mode:
                slidemaker.ImgTextSlideMaker(node, prs_info, prs)
            else:
                slidemaker.TextSlideMaker(node, prs_info, prs)
    
    # If the node's layer is 2 and the level_mode is true, makes a second level section slide
    # At the same time, if the node's note is not empty, make a pure text slide or text with an image slide following it.
    # Otherwise, if the level_mode is false, treat it like a normal node.
    
    elif node.layer == 2:
        
        if level_mode:
        
            slidemaker.SectionSlideMaker(node, prs_info, prs, 2)
            
            if node.note != '':
                if img_mode:
                    slidemaker.ImgTextSlideMaker(node, prs_info, prs)
                else:
                    slidemaker.TextSlideMaker(node, prs_info, prs)
                    
        else:
            
            if img_mode:
                slidemaker.ImgTextSlideMaker(node, prs_info, prs)
            else:
                slidemaker.TextSlideMaker(node, prs_info, prs)
            
    # Otherwise, makes a pure text slide or text with image slide.
            
    else:
        
        if img_mode:
            slidemaker.ImgTextSlideMaker(node, prs_info, prs)
        else:
            slidemaker.TextSlideMaker(node, prs_info, prs)

    # Traversally calls this function for all the children of the node.

    for child in node.child:
        standard_strategy(child, prs_info, prs, level_mode, img_mode)


def guide_through(loc_in, loc_out):
    """This function provides an interactive way for the user to use AutoPre.
    
    And it now serves as the default way of exploiting the capacity of AutoPre.
    
    Args:
        loc_in: A string that indicates folder that contains the content files (.opml or .docx).
        loc_out: A string that indicates folder that contains the output files (.pptx).
    """
    
    # Introduction.
    
    print("AutoPre:\n"
          "  Thank you for using AutoPre, a PPT automatic filler based on Dynalist note, Word document and so on, "
          "targeting on freeing ourselves from the exhausting ppt formating. "
          "Thus, we can focus on the content itself and contribute more valuable outputs.\n")
    
    print("  感谢您使用AutoPre，一款基于Dynalist笔记、Word文档等的ppt自动填充软件，旨在让我们从繁重的ppt制作中解脱出来. "
          "这样，我们就可以专注于内容本身，贡献更多有价值的产出.\n")
    
    # Files loading.
    
    print("AutoPre:")
    
    root_nodes, name_list = load.load_files(loc_in)
    
    print("\n  {} files has been loaded.".format(len(root_nodes)))
    
    print("  {}份文件被加载.\n".format(len(root_nodes)))
    
    # Setting and Outputting.
    
    for n in range(len(root_nodes)):
        
        # Sets the output parameters.
        
        print("AutoPre:\n  Setting the {}'s output parameters.".format(name_list[n]))
        
        print("  设置 {} 的输出参数.\n".format(name_list[n]))
        
        prs_info = {
            'title': None,
            'date': None,
            'part_num': 0,
            'part_title': None,
            'color': None,
            'author': None
            }
        
        # Sets the title of the presentation.
        
        print("  Please enter the name of this presentation.", end = '')
        
        prs_info['title'] = input("  请输入该演示文档的标题.\n  ")
        
        # Sets the date of the presentation.
        
        print("\n  Please enter the date statement of presentation. eg.Date  7/11/1917", end = '')
        
        prs_info['date'] = input("  请输入演示文档的日期说明. 示例:日期  1917/11/7\n  ")
        
        # Sets the author of the presentation.
        
        print("\n  Please enter the inscribe includes the author's name. eg.Reporter Kirov", end = '')
        
        prs_info['author'] = input("  请输入包含作者姓名的落款. 示例:汇报  李明\n  ")
        
        # Sets the template of the presentation.
          
        names = ''
        for name in THEME_AND_COLORS.keys():
            names = names + name + '  '
        
        print("\n  Please enter the template name. Here are the choices:", end = '')
        
        template = input("  请输入模板名称，以下为选择范围:\n  {}\n  ".format(names))
            
        if template not in THEME_AND_COLORS.keys():
            
            print("\n  Input is invalid. Please note the case. Now using default template Marxist.")
            
            print("  非法输入，请注意大小写. 采用默认模板 Marxist.")
            
            template = 'Marxist'
        
        # Sets the themecolor according to the template used.
        
        prs_info['color'] = THEME_AND_COLORS[template]
        
        # Defines whether the second level sections pages will be used.
            
        print("\n  Would you like to use second level sections pages?\n"
              "  Please enter 'Y' representing yes or 'N' representing no.", end = '')
        
        level_mode = input("  您希望使用二级章节分界页面吗？\n"
                         "  请输入 Y 代表是，输入 N 代表否.\n  ")
        
        if level_mode not in ['Y', 'N']:
            
            print("\n  Input is invalid. Please note the case. Now using default setting, i.e. uses image illustration.")
            
            print("  非法输入，请注意大小写. 采用默认选项即使用二级章节分界页面.")
            
            level_mode = 'Y'
        
        if level_mode == 'Y':
            level_mode = True
            
        else:
            level_mode = False
        
        # Defines whether the image illustration will be used.
            
        print("\n  Would you like to add image illustration to each main body pages?\n"
              "  Please enter 'Y' representing yes or 'N' representing no.", end = '')
        
        img_mode = input("  您希望给正文页面添加插图区吗？\n"
                         "  请输入 Y 代表是，输入 N 代表否.\n  ")
        
        if img_mode not in ['Y', 'N']:
            
            print("\n  Input is invalid. Please note the case. Now using default setting, i.e. don't use image illustration.")
            
            print("  非法输入，请注意大小写. 采用默认选项即正文页面不使用插图.")
            
            img_mode = 'N'
        
        if img_mode == 'Y':
            img_mode = True
            
        else:
            img_mode = False
        
        # Outputs the presentation file.
        
        prs = Presentation('templates/' + template + '.pptx')
        
        standard_strategy(root_nodes[n], prs_info, prs, level_mode, img_mode)
        slidemaker.BackCoverSlideMaker(prs)
        
        prs.save(loc_out + prs_info['title'] + '.pptx')
        
        print("\n  Presentation document is outputted to " + loc_out + prs_info['title'] + '.pptx')
        
        print("  演示文档输出到 " + loc_out + prs_info['title'] + '.pptx\n')
    

if __name__ == '__main__':
    
    guide_through("../documents/", "../outputs/")
    
    