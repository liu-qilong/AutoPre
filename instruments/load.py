import re
import os
import docx

class OutlineNode(object):
    """Class of outline document's node. 
    
    The outline documents are organized with the basic element, node, in the architecture of nodes' tree.'
    
    Attributes:
        title: A string contains the title of the node.
        note: A String contains the note content of the node.
        layer: An integer count of the layer of the node.
        child: A list of OutlineNode objects that are children of this node.
        father: A OutlineNode object that is the father of this node.
    """
    
    def __init__(self, title, note, layer):
        
        self.title = self.content_tidy(title)
        self.note = self.content_tidy(note)
        self.layer = int(layer)
        
        self.child = []
        self.father = None

    def content_tidy(self, content):
        """Execute content replacement to tidy it up."""
        
        content = content.replace('&#10;', '\n')
        content = content.replace('&quot;', '"')
        content = content.replace('$$^i$$', '')
        content = content.replace('&lt;', '<')
        content = content.replace('&gt;', '>')
        content = content.replace('&apos;', "'")
        content = content.replace('&amp;', "&")
        content = content.replace('?', '')
        
        return content

    def set_father(self, fatherSet):
        
        self.father = fatherSet

    def add_child(self, childSet):
        
        self.child.append(childSet)

    def have_code(self):
        
        if '`' in self.note:
            return True
        else:
            return False

    def have_formula(self):
        
        if '$$' in self.note:
            return True
        else:
            return False

    def __str__(self):
        title = ' '*self.layer*2 + self.title + '\n'
        note = ' '*self.layer*2 + self.note
        return title + note
    
    def traversal_print(self):
        """Traversally prints the whole nodes' tree."""

        print(self)
        
        for child in self.child:
            child.traversal_print()
            
            
def read_opml(location):
    """Reads the opml file and analyses it to build a nodes' tree of its content."""
    
    def analyse_opml(opml_content):
        """Analyses the opml codes and builds a nodes' tree for it.
        
        Noted that the basic element of the nodes' tree is OutlineNode.
        
        Attributes:
            opml_content: A string contains the whole content of the opml code.
            
        Returns:
            The root node of the nodes' tree.
        """
        
        # The content of the opml code is split and analysed line by line.
        
        opml_content = re.split('\n', opml_content)
    
        for content in opml_content:
            
            # Only the lines that contain "outline" or "text" are valuable.
            
            if 'outline' not in content:
                continue
            
            if 'text' not in content:
                continue
    
            # Split the content of the line again which creates a list contains: [spaces, title, note, useless content]
            # The number of spaces is related to the layer of the node: layer = number of spaces / 2 - 2
            # Then creates an outline node object.
            
            node_split = re.split('<outline text="|" _note="|" heading="1"|" heading="2"|" heading="3"|"/>|">|>',content)
            layer = len(node_split[0])/2 - 2
            new_node = OutlineNode(node_split[1], node_split[2], layer)
    
            # If the new node is in layer 0, then it should be the root node.
            # Saves it as the root node and current node and turns to the next loop.
            
            if layer == 0:
                root_node = new_node
                current_node = new_node
                continue
            
            # If the layer of the new node is larger than the current node, then it should be the current node's child.
            
            if layer > current_node.layer:
                new_node.set_father(current_node)
                current_node.add_child(new_node)
    
            # If the layer of the new node is equal to the current node, then they should share the same father.
    
            if layer == current_node.layer:
                new_node.set_father(current_node.father)
                current_node.father.add_child(new_node)
            
            # If the layer of the new node is smaller than the current node, then find the node, from the current node's fathers serial,
            # in the same layer with the new node. They should share the same father.
            
            if layer < current_node.layer:
                gap = current_node.layer - layer
                
                while gap > 0:
                    current_node = current_node.father
                    gap -= 1
                    
                new_node.set_father(current_node.father)
                current_node.father.add_child(new_node)
            
            # Refresh the current node.
            
            current_node = new_node
        
        return root_node

    with open (location, 'r', encoding='UTF-8') as file:
        opml_content = file.read()

    root_node = analyse_opml(opml_content)
    return root_node


def read_docx(location):
    """Reads the opml file and analyses it to build a nodes' tree of its content."""
    
    def sorting_docx(location):
        """Sorting the docx content for the following analysing.
        
        The main targets are distinguishing the titles and the corresponding notes and the layer they belong to.
        
        Attributes:
            location: The string that contains the location of the docx file. 
            
        Returns:
            A list of the OutlineNodes.
        """
        
        # Loads the whole content of the docx file into a docx document object.
        # Distinguishs the titles and the corresponding notes and the layer they belong to.
        
        titles = []
        notes  = []
        layers = []
        
        docx_content = docx.Document(location)
        
        for paragraph in docx_content.paragraphs:
            
            # If the style of the paragraph is not "Normal" then it is a title.
            
            if paragraph.style != docx_content.styles['Normal']:
                titles.append(paragraph.text)
                notes.append('')
                
                # Matchs the layer of the tiele.
                
                if paragraph.style == docx_content.styles['Heading 1']:
                    layers.append(1)
                
                if paragraph.style == docx_content.styles['Heading 2']:
                    layers.append(2)
                
                if paragraph.style == docx_content.styles['Heading 3']:
                    layers.append(3)
            
            # Otherwise, it's a part of the note.
            
            else:
                notes[len(titles) - 1] = notes[len(titles) - 1] + paragraph.text + '\n'
        
        # Creates a list that contains the nodes and returns it.
        
        nodes = []        
        for n in range(len(titles)):
            node = OutlineNode(titles[n], notes[n], layers[n])
            nodes.append(node)
            
        return nodes
            
    
    def analyse_docx(nodes):
        """Analyses the nodes' structure and builds a nodes' tree for it.
        
        Noted that the basic element of the nodes' tree is OutlineNode.
        
        Attributes:
            nodes: A list of the OutlineNodes.
            
        Returns:
            The root node of the nodes' tree.
        """
        
        root_node = OutlineNode('', '', 0)
        current_node = root_node
        
        for new_node in nodes:
            
            # If the layer of the new node is larger than the current node, then it should be the current node's child.
            
            if new_node.layer > current_node.layer:
                new_node.set_father(current_node)
                current_node.add_child(new_node)
    
            # If the layer of the new node is equal to the current node, then they should share the same father.
    
            if new_node.layer == current_node.layer:
                new_node.set_father(current_node.father)
                current_node.father.add_child(new_node)
            
            # If the layer of the new node is smaller than the current node, then find the node, from the current node's fathers serial,
            # in the same layer with the new node. They should share the same father.
            
            if new_node.layer < current_node.layer:
                gap = current_node.layer - new_node.layer
                
                while gap > 0:
                    current_node = current_node.father
                    gap -= 1
                    
                new_node.set_father(current_node.father)
                current_node.father.add_child(new_node)
            
            # Refresh the current node.
            
            current_node = new_node
        
        return root_node
    
    nodes = sorting_docx(location)
    rootNode = analyse_docx(nodes)
    return rootNode


def load_files(location):
    """Loads all the valid files from an indicated folder.
    
    It analyses all the valid files from the provided location and creates OutlineNodes' tree for each one.'
    
    Args:
        location: A string that contains the folder location of the target file.
        
    Returns:
        A list that contains the root nodes.
        A list of the file names."""

    file_list = os.listdir(location)
    
    # Loads the files according to its format.
    
    root_nodes = []
    name_list = []
    for file in file_list:
        
        if '.opml' in file:
            
            root_nodes.append(read_opml(location + file))
            name_list.append(file)
            print("  {} has been loaded.".format(file))
            
        elif '.docx' in file:
            
            if '~$' in file: # Which indicates that this is a temporary file of no value.
                break
            
            root_nodes.append(read_docx(location + file))
            name_list.append(file)
            print("  {} has been loaded.".format(file))
            
        else:
            
            pass
    
    return root_nodes, name_list


# Functions testing.

if __name__ == '__main__':
    
    location = "../documents/"
    root_nodes, name_list = load_files(location)
    root_nodes[0].traversal_print()
    
    