import meraki
from rich import print as pp
from rich.console import Console
from rich.table import Table
from openpyxl import Workbook
from openpyxl.styles import Font, Color

def select_org():
    # Fetch and select the organization
    print('\n\nFetching organizations...\n')
    organizations = dashboard.organizations.getOrganizations()
    ids = []
    table = Table(title="Meraki Organizations")
    table.add_column("Organization #", justify="left", style="cyan", no_wrap=True)
    table.add_column("Org Name", justify="left", style="cyan", no_wrap=True)
    counter = 0
    for organization in organizations:
        ids.append(organization['id'])
        table.add_row(str(counter), organization['name'])
        counter+=1
    console = Console()
    console.print(table)
    isOrgDone = False
    while isOrgDone == False:
        selected = input('\nKindly select the organization ID you would like to query: ')
        try:
            if int(selected) in range(0,counter):
                isOrgDone = True
            else:
                print('\t[red]Invalid Organization Number\n')
        except:
            print('\t[red]Invalid Organization Number\n')
    return(organizations[int(selected)]['id'], organizations[int(selected)]['name'])


if __name__ == '__main__':
    # Initializing Meraki SDK
    dashboard = meraki.DashboardAPI(output_log=False, suppress_logging=True)
    org_id, org_name = select_org()
    results = {}
    #
    # Get networks
    #
    networks = dashboard.organizations.getOrganizationNetworks(org_id)
    uplinks = dashboard.organizations.getOrganizationUplinksStatuses(org_id)
    for network in networks:
        network_id = network['id']
        network_name = network['name']
        results[network_id] = {'name': network_name, 'publicIp': ""}
    for uplink in uplinks:
        if len(uplink['uplinks']) > 0:
            results[uplink['networkId']]['publicIp'] = uplink['uplinks'][0]['publicIp']
    # 
    # Create Excel file
    #
    workbook = Workbook()
    sheet = workbook.active
    #
    sheet["A1"] = "Network Name"
    sheet["B1"] = "Public IP"
    line = 2
    #
    for network_id in results:
        sheet["A"+str(line)] = results[network_id]['name']
        sheet["B"+str(line)] = results[network_id]['publicIp']
        line += 1
    #
    # Save Excel file
    #
    workbook.save(f"{org_name} public ip report.xlsx")
    pp("\n\n\t[green]Excel file created successfully\n\n")