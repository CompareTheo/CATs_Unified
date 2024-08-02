import re

class ParseLine:
    def __init__(self, order_line, headers):
        self.headers = headers
        
        # Skip null lines
        if "\t" not in order_line:
            self.line_is_null = True
        else:
            self.line_is_null = False
        
            # Parse
            fields = order_line.split("\t")
            
            # Title
            self.title  = fields[0]
            self.title_clean = re.sub('[,:;."]', '', self.title)
            self.title_clean = re.sub('[-]', ' ', self.title_clean)
            self.title_clean = re.sub('[&]', ' ', self.title_clean)
            self.title_split = self.title_clean.split(" ")
            
            self.title_parsed_array = []
            loop_counter = 0
            for t in self.title_split:
                if loop_counter < 5:
                    self.title_parsed_array.append(t)
                    loop_counter += 1
            self.title_parsed = " ".join(self.title_parsed_array)
            
            self.title_short = self.title_split[0]
            
            # Author
            self.author = fields[6]
            a = self.author.split(", ")
            self.author_lastname = a[0]
            
            # Editor
            self.editor = fields[7]
            e = self.editor.split(", ")
            
            # If no author, switch to editor
            if self.author == "":
                self.author = e[0]
                self.author_lastname = e[0]
            
            # Keywords
            self.kw = f"{self.author_lastname} {self.title_parsed}"
            if self.author_lastname == "":
                self.kw = f"{self.title_parsed}"
            
            # ISBN
            self.isbn   = fields[10]
            
            # Publisher
            self.pub    = fields[8]
            p = self.pub.split(" ")
            self.pub_short = p[0]
            
            # Pub year
            self.pub_year = fields[9]
            
            # Binding
            self.binding = fields[11]
            if self.binding.lower() != "ebook":
                self.line_is_null = True
                return

            # Selector
            self.selector = fields[167]  # This index may need adjustment based on your actual data
            
            # Duplication note
            self.intdup = fields[170]  # This index may need adjustment based on your actual data
            if not self.intdup or len(self.intdup.strip()) == 0:
                self.dupe_is_null = True
            else:
                self.dupe_is_null = False

            # Purchase option
            self.purchase_option = self.get_purchase_option(fields)
    
    def get_purchase_option(self, fields):
        purchase_options = []
        purchase_types = ["Non-Linear", "Unlimited Access", "Concurrent Access"]
        for i, header in enumerate(self.headers):
            if header.endswith('.Purchase_Option'):
                option = fields[i]
                if any(purchase_type in option for purchase_type in purchase_types):
                    purchase_options.append(option)
        return purchase_options