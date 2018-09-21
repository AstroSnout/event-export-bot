import xlsxwriter
import datetime


class Spreadsheets:
    def __init__(self):
        # Cell formats - class name
        dk_cell = {'font_name': 'Arial', 'font_size': 10, 'bg_color': '#C41F3B', 'bold': True}
        dh_cell = {'font_name': 'Arial', 'font_size': 10, 'bg_color': '#A330C9', 'bold': True}
        druid_cell = {'font_name': 'Arial', 'font_size': 10, 'bg_color': '#FF7D0A', 'bold': True}
        hunter_cell = {'font_name': 'Arial', 'font_size': 10, 'bg_color': '#ABD473', 'bold': True}
        mage_cell = {'font_name': 'Arial', 'font_size': 10, 'bg_color': '#40C7EB', 'bold': True}
        monk_cell = {'font_name': 'Arial', 'font_size': 10, 'bg_color': '#00FF96', 'bold': True}
        paladin_cell = {'font_name': 'Arial', 'font_size': 10, 'bg_color': '#F58CBA', 'bold': True}
        priest_cell = {'font_name': 'Arial', 'font_size': 10, 'bg_color': '#FFFFFF', 'bold': True}
        rogue_cell = {'font_name': 'Arial', 'font_size': 10, 'bg_color': '#FFF569', 'bold': True}
        shaman_cell = {'font_name': 'Arial', 'font_size': 10, 'bg_color': '#0070DE', 'bold': True}
        warlock_cell = {'font_name': 'Arial', 'font_size': 10, 'bg_color': '#8787ED', 'bold': True}
        warr_cell = {'font_name': 'Arial', 'font_size': 10, 'bg_color': '#C79C6E', 'bold': True}
        # Cell formats - invite status
        accepted_cell = {'font_name': 'Arial', 'font_size': 10, 'bg_color': 'green', 'bold': True}
        standby_cell = {'font_name': 'Arial', 'font_size': 10, 'bg_color': 'cyan', 'bold': True}
        tentative_cell = {'font_name': 'Arial', 'font_size': 10, 'bg_color': 'yellow', 'bold': True}
        decline_cell = {'font_name': 'Arial', 'font_size': 10, 'bg_color': 'red', 'bold': True}
        # Cell Type to Cell Format translator
        self.cell_format = {
            'Death Knight': dk_cell,
            'Demon Hunter': dh_cell,
            'Druid': druid_cell,
            'Hunter': hunter_cell,
            'Mage': mage_cell,
            'Monk': monk_cell,
            'Paladin': paladin_cell,
            'Priest': priest_cell,
            'Rogue': rogue_cell,
            'Shaman': shaman_cell,
            'Warlock': warlock_cell,
            'Warrior': warr_cell,
            'Invited': tentative_cell,
            'Accepted': accepted_cell,
            'Declined': decline_cell,
            'Confirmed': accepted_cell,
            'Out': decline_cell,
            'Standby': standby_cell,
            'Signed Up': accepted_cell,
            'Not Signed Up': decline_cell,
            'Tentative': tentative_cell
        }
        self.invite_status = {
                1: 'Invited',
                2: 'Accepted',
                3: 'Declined',
                4: 'Confirmed',
                5: 'Out',
                6: 'Standby',
                7: 'Signed Up',
                8: 'Not Signed Up',
                9: 'Tentative'
            }

    def make_spreadsheet(self, data_type, data):
        if data_type == '0':
            event_title = data['eventInfo']['title']
            event_date = data['eventInfo']['eventDate']
            del data['eventInfo']

            # Making of the spreadsheet
            for char in '\\/:*?<>|':
                event_title = event_title.replace(char, '-')
            print(event_title)
            wb_name = f'{event_date} {event_title}.xlsx'
            wb = xlsxwriter.Workbook(wb_name, {'in_memory': True})
            ws = wb.add_worksheet()

            # Widen columns in range
            ws.set_column('A:A', 10)
            ws.set_column('B:C', 40)
            ws.set_column('D:D', 10)

            header = wb.add_format({'bold': True, 'font_size': 24, 'center_across': True, 'border': 2})
            fill = wb.add_format({'fg_color': '#000000'})

            ws.write(f'A1', ' ', fill)
            ws.write(f'B1', ' ', fill)
            ws.write(f'C1', ' ', fill)
            ws.write(f'D1', ' ', fill)

            ws.write('A2', ' ', fill)
            ws.write('B2', 'Character Name', header)
            ws.write('C2', 'Invite Status', header)
            ws.write('D2', ' ', fill)

            # Populate the rows
            for i in range(1, len(data)+1):
                i = str(i)
                # Assign variables for increased readability
                invite_status = self.invite_status[data[i]['stat']]
                class_name = data[i]['cls']
                char_name = data[i]['name']
                # Cell value set to character's name, cell style set to character's class (class color cell BG for now only)
                class_cell = wb.add_format(self.cell_format[class_name])
                invite_cell = wb.add_format(self.cell_format[invite_status])
                i = int(i)
                ws.write(f'A{str(i+2)}', ' ', fill)
                ws.write(f'B{str(i+2)}', char_name, class_cell)
                ws.write(f'C{str(i+2)}', invite_status, invite_cell)
                ws.write(f'D{str(i+2)}', ' ', fill)

            # Close the sheet
            ws.write(f'A{len(data)+3}', ' ', fill)
            ws.write(f'B{len(data)+3}', ' ', fill)
            ws.write(f'C{len(data)+3}', ' ', fill)
            ws.write(f'D{len(data)+3}', ' ', fill)

            wb.close()

        elif data_type == '1':
            print(data)
            all_online = [*data]

            # Making of the spreadsheet
            wb_name = f'{datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")}_Online_people.xlsx'
            wb = xlsxwriter.Workbook(wb_name, {'in_memory': True})
            ws = wb.add_worksheet()

            # Widen columns in range
            ws.set_column('A:A', 10)
            ws.set_column('B:B', 40)
            ws.set_column('C:C', 10)

            header = wb.add_format({'bold': True, 'font_size': 24, 'center_across': True, 'border': 2})
            body = wb.add_format({'font_size': 14, 'border': 1})
            fill = wb.add_format({'fg_color': '#000000'})

            ws.write(f'A1', ' ', fill)
            ws.write(f'B1', ' ', fill)
            ws.write(f'C1', ' ', fill)

            ws.write('A2', ' ', fill)
            ws.write('B2', 'Online Characters', header)
            ws.write('C2', ' ', fill)

            # Populate the rows
            for i in range(len(all_online)):
                char_name = all_online[i]

                ws.write(f'A{i+3}', ' ', fill)
                ws.write(f'B{i+3}', char_name, body)
                ws.write(f'C{i+3}', ' ', fill)

            # Close the sheet
            ws.write(f'A{len(all_online)+3}', ' ', fill)
            ws.write(f'B{len(all_online)+3}', ' ', fill)
            ws.write(f'C{len(all_online)+3}', ' ', fill)

            wb.close()

        elif data_type == '2':
            # Making of the spreadsheet
            wb_name = f'{datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")}_All_members.xlsx'
            wb = xlsxwriter.Workbook(wb_name, {'in_memory': True})
            ws = wb.add_worksheet()

            # Widen columns in range
            ws.set_column('A:A', 10)
            ws.set_column('B:D', 40)
            ws.set_column('E:E', 10)

            header = wb.add_format({'bold': True, 'font_size': 24, 'center_across': True, 'border': 2})
            body = wb.add_format({'font_size': 14, 'border': 1})
            fill = wb.add_format({'fg_color': '#000000'})

            ws.write(f'A1', ' ', fill)
            ws.write(f'B1', ' ', fill)
            ws.write(f'C1', ' ', fill)
            ws.write(f'D1', ' ', fill)
            ws.write(f'E1', ' ', fill)

            ws.write('A2', ' ', fill)
            ws.write('B2', 'Character Name', header)
            ws.write('C2', 'Member Note', header)
            ws.write('D2', 'Officer Note', header)
            ws.write('E2', ' ', fill)

            # Populate the rows
            for i in range(1, len(data)+1):
                i = str(i)
                char_name = data[i]['name']
                try:
                    member_note = data[i]['memberNote']
                except KeyError:
                    member_note = 'N/A'
                try:
                    officer_note = data[i]['officerNote']
                except KeyError:
                    officer_note = 'N/A'

                i = int(i)
                # Cell value set to character's name, cell style set to character's class (class color cell BG for now only)
                ws.write(f'A{str(i+2)}', ' ', fill)
                ws.write(f'B{str(i+2)}', char_name, body)
                ws.write(f'C{str(i+2)}', member_note, body)
                ws.write(f'D{str(i+2)}', officer_note, body)
                ws.write(f'E{str(i+2)}', ' ', fill)

            ws.write(f'A{len(data)+3}', ' ', fill)
            ws.write(f'B{len(data)+3}', ' ', fill)
            ws.write(f'C{len(data)+3}', ' ', fill)
            ws.write(f'D{len(data)+3}', ' ', fill)
            ws.write(f'E{len(data)+3}', ' ', fill)

            # Close the sheet
            wb.close()

        elif data_type == '3':
            print(data)
            all_in_raid = [*data]

            # Making of the spreadsheet
            wb_name = f'{datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")}_People_in_raid.xlsx'
            wb = xlsxwriter.Workbook(wb_name, {'in_memory': True})
            ws = wb.add_worksheet()

            # Widen columns in range
            ws.set_column('A:A', 10)
            ws.set_column('B:B', 40)
            ws.set_column('C:C', 10)

            header = wb.add_format({'bold': True, 'font_size': 24, 'center_across': True, 'border': 2})
            body = wb.add_format({'font_size': 14, 'border': 1})
            fill = wb.add_format({'fg_color': '#000000'})

            ws.write(f'A1', ' ', fill)
            ws.write(f'B1', ' ', fill)
            ws.write(f'C1', ' ', fill)

            ws.write('A2', ' ', fill)
            ws.write('B2', 'Characters in raid', header)
            ws.write('C2', ' ', fill)

            # Populate the rows
            for i in range(len(all_in_raid)):
                char_name = all_in_raid[i]

                ws.write(f'A{i+3}', ' ', fill)
                ws.write(f'B{i+3}', char_name, body)
                ws.write(f'C{i+3}', ' ', fill)

            # Close the sheet
            ws.write(f'A{len(all_in_raid)+3}', ' ', fill)
            ws.write(f'B{len(all_in_raid)+3}', ' ', fill)
            ws.write(f'C{len(all_in_raid)+3}', ' ', fill)

            wb.close()
        else:
            raise Exception

        return wb_name
