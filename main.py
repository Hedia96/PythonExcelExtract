from Utils.excelreader import *
import argparse

def main ():
    configpath ="Config/config.yaml"
    config = load_config(configpath)
    spreadsheet_name =config['spreadsheet_name']  
    
     
if __name__ == "__main__":
    main()
