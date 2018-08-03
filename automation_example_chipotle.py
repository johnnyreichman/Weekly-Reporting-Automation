#This program optimizes the Chipotle weekly reporting process. 
#Now it takes ~1'20" to process 6 files

from run_macro import *
from convert_csv import *
from vba_extract import *

def runMacroAndConvertToCSV(input_file, macro_name, macro_file, output_file = "default_output"):
	print("\nProcessing: ", input_file)
	#checks if the macro file given has had it's vba code extracted or not
	if(macro_file[-4:]) != ".bin":
		macro_file = extractVBA(macro_file)
	addMacro(input_file, macro_file, output_file)
	runMacro(input_file, macro_name, output_file)
	exportCSV(output_file + ".xlsm", output_file + ".csv")


def main():
	runMacroAndConvertToCSV("FB+IG All Campaigns 6.18-7.1.xlsx", "Macro_FB_IG_All_Campaigns", "MACRO_FB_IG_All Campaigns.xlsm", "FB_IG_All_Campaigns_processed")
	runMacroAndConvertToCSV("IAS_2018-05-21.xlsx", "IAS", "IAS.xlsm", "IAS_processed")
	runMacroAndConvertToCSV("Sales 6-25 to7-1.xlsx", "Display_AllCampaigns", "Display_AllCampaigns.xlsm", "Display_AllCampaigns_processed")
	runMacroAndConvertToCSV("Chipotle Keywords SW & Queso.xlsx", "Search_Keywords_Macro", "SEARCH_ Keywords SW Queso_MACRO.xlsm", "Search_Keywords_processed")
	runMacroAndConvertToCSV("Chipotle Hour of Day SW & Queso.xlsx", "Chipotle_Search_HourofDay", "PaidSearch_HourOfDay.xlsm", "Search_HourOfDay_processed")
	runMacroAndConvertToCSV("FB+IG DigComm 6.18-7.1.xlsx", "MACRO_FB_IG_DigCom", "MACRO_FB+IG_Digcom.xlsm", "FB_IG_DigCom_processed")




if __name__ == "__main__":
	main()
