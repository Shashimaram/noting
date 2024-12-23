from utilities import writing_to_file
import config

class VPC_cost_assessment():
    def __init__(self,workbook):
        self.workbook = workbook

    




    def process(self):
        counter=1
        for row in range(self.workbook.max_row+1):
            counter+=1
            cell_obj = self.workbook.cell(row = counter,column=config.usage_amount_column)
            if cell_obj.value != None:
                usagetype = cell_obj.value
                usagetype_tolist = usagetype.lower().split('-')

                if "vpcendpoint" in usagetype_tolist:
                    writing_to_file(workbook=self.workbook,
                                    row=counter,
                                    column_10_value=0,
                                    column_11_value=0,
                                    column_12_value="OCI private endpoints",
                                    column_13_value="OCI private endpoints is available for no additional charge. There is no per-hour connection fee or per-byte data processing fee,")

                elif "transitgateway" in usagetype_tolist:
                    writing_to_file(workbook=self.workbook,
                                    row=counter,
                                    column_10_value=0,
                                    column_11_value=0,
                                    column_12_value="Dynamic Routing Gateway (DRG)",
                                    column_13_value="OCI Dynamic Routing Gateway is available for no additional charge. There is no per-hour connection fee or per-byte data processing fee,")
                

                elif "publicipv4:inuseaddress" in usagetype_tolist or "publicipv4:idleaddress" in usagetype_tolist:
                    writing_to_file(workbook=self.workbook,
                                    row=counter,
                                    column_10_value=0,
                                    column_11_value=0,
                                    column_12_value="Reserved IP4",
                                    column_13_value="No additional charge for in-use or idle address")
                
                elif 'vpn' in usagetype_tolist:
                    writing_to_file(workbook=self.workbook,
                                    row=counter,
                                    column_10_value=0,
                                    column_11_value=0,
                                    column_12_value="OCI site-to-site VPN endpoints",
                                    column_13_value="There is no per-hour connection fee or per-byte data processing fee,")

                
                else:
                    print(usagetype_tolist)
