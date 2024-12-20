#!/usr/bin/python3
import concurrent.futures
import requests
import xmltodict

__version__ = '0.1.2'
__api_version__ = '1.2'

# main class ##################################################################
class CollectionCheck:
    def __init__(self, temporary_collection_list):
        self.collection_list = temporary_collection_list
    
    def check_collection(self, collection_id):
        return collection_id in self.collection_list

class SRU():
    def __init__(self, r, zone="", inst_code="", sru_path=""):
        self.zone = zone
        self.r = r
        self.inst_code = inst_code
        self.xml = r.text or ""
        self.dict = xmltodict.parse(self.xml, dict_constructor=dict)
        
        self.print_holdings = []
        self.e_holdings = []
        
        self.call_number = ""
        self.location = ""
        
        # get number of records, check for errors
        try:
            self.numberOfRecords = int(self.dict['searchRetrieveResponse']['numberOfRecords'])
            self.ok = True
            self.errors = None
        except Exception as e:
            self.numberOfRecords = 0
            self.ok = False
            self.errors = self.dict['searchRetrieveResponse']['diagnostics']['diag:diagnostic']['diag:message']
            print(f"{e}\n{self.xml}")
            
        # get print holdings
        if self.numberOfRecords > 0 and zone == "IZ":
            try:
                self.records = self.dict['searchRetrieveResponse']['records']['record']
                self.print_holdings, self.location, self.call_number = get_print_holdings(self.records)
            except Exception as e:
                self.print_holdings = []
                print(f"{e}\n\n{self.xml}")
                
        # get e-holdings
        if self.numberOfRecords > 0:
            try:
                self.records = self.dict['searchRetrieveResponse']['records']['record']
                self.e_holdings = get_e_holdings(self.records, zone=self.zone, inst_code = self.inst_code)
            except Exception as e:
                self.e_holdings = []
                print(f"{e}\n\n{self.xml}")
                
        # set e-availability
        if self.e_holdings != []:
            self.have_e_holdings = True
        else:
            self.have_e_holdings = False
            
# search functions ############################################################
def search(query=""):
    r = requests.get(query)
    return r
    
def load_url(url):
    r = requests.get(url)
    return r
    
def searches(urls, workers):
    with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as executor:
        r_array = executor.map(load_url, urls)
                
    return r_array
    
def get_base_url(zone):
    return f"{ENDPOINTS[zone]}"
    
def get_query_url(query, operation="searchRetrieve", recordSchema="marcxml", maximumRecords="10", startRecord="1"):
    return(f"?version={__api_version__}&operation={operation}&recordSchema={recordSchema}&maximumRecords={maximumRecords}&startRecord={startRecord}&query={query}")

def make_url(zone="", sru_path="", query=""):
    url = f"{sru_path}{get_query_url(query)}"
    return (url)
    
# parse functions #############################################################
def parse(r, zone="", inst_code=""):
    sru_object = SRU(r, zone, inst_code)
    return sru_object
    
def get_print_holdings(records):
    print_holdings = []
    
    # parse SRU response
    for record in records:
        try:
            datafields = record['recordData']['record']['datafield']
        except Exception as e:
            #print(e)
            datafields = records['recordData']['record']['datafield']
            
        for field in datafields:
            code_c = ""
            code_d = ""
            range = ""
            code_t = []
            # Check for print holdings
            if field['@tag'] == "AVA":
                for subfield in field['subfield']:
                    if subfield['@code'] == '8':
                        code_8 = subfield['#text']
                    if subfield['@code'] == 'c':
                        code_c = subfield['#text']
                    if subfield['@code'] == 'd':
                        code_d = subfield['#text']
                    if subfield['@code'] == 'e':
                        code_e = subfield['#text']
                    if subfield['@code'] == 'm':
                        code_m = subfield['#text']
                    if subfield['@code'] == 's':
                        code_s = subfield['#text']
                    if subfield['@code'] == 't':
                        #code_t = subfield['#text']
                        code_t.append(subfield['#text'])
                
                # Join code_t string
                range = "\n".join(code_t)
                
                # Check for print holdings
                if code_e == "available" or code_e == "Available":
                    print_holdings_statement = f"{range} ({code_c})"
                    if print_holdings_statement not in print_holdings:
                        print_holdings.append(print_holdings_statement)
                
    return print_holdings, code_c, code_d
    
def get_e_holdings(records, zone="", inst_code="", match_list=None):
    e_holdings = []
    cats_ready = False  # Track if there's a match for CATs Ready
    subfield_0_value = None  # Store the Subfield 0 value if a match is found

    if not records:
        return e_holdings, cats_ready, subfield_0_value  # Return defaults if no records

    # Parse SRU response
    for record in records:
        try:
            datafields = record['recordData']['record']['datafield']
        except Exception:
            datafields = records['recordData']['record']['datafield']
            
        for field in datafields:
            code_c = ""
            code_e = ""
            code_m = ""
            code_0 = ""  # Subfield 0 value
        
            # Check for electronic access
            if field['@tag'] == "AVE":
                for subfield in field['subfield']:
                    if subfield['@code'] == 'c':
                        code_c = subfield['#text']
                    if subfield['@code'] == 'e':
                        code_e = subfield['#text']
                    if subfield['@code'] == 'm':
                        code_m = subfield['#text']
                    if subfield['@code'] == '0':
                        code_0 = subfield['#text']

                # Check if the code_c matches any value in match_list
                if match_list and code_c in match_list:
                    cats_ready = True
                    subfield_0_value = code_0  # Capture the Subfield 0 value
                
                # Check for e-holdings in IZ
                if zone == "IZ" and code_e == "Available":
                    e_holdings_statement = f"{code_m} ({code_c})"
                    if e_holdings_statement not in e_holdings:
                        e_holdings.append(e_holdings_statement)
                
    return e_holdings, cats_ready, subfield_0_value