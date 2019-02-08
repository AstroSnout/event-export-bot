import discord
from discord.ext import commands

import sys
import traceback
import base64 as b64
import json
import datetime
import os
import xlsxwriter
import aiohttp
import yarl
import requests

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
cell_format = {
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
invite_status = {
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



BNET_API_KEY = '4f50adccbaad4636872eac30e7d5f4d5'
BNET_API_SECRET = 'Bu3kf5npXVZhvvOYiL7rX9820q5jJ72d'

access_token_grant = f'https://eu.battle.net/oauth/token?grant_type=client_credentials&client_id={BNET_API_KEY}&client_secret={BNET_API_SECRET}'

bnet_token = json.loads(
    requests.get(access_token_grant).content
)['access_token']

async def get_json(uri, timeout=60):
    print('Requesting JSON ->', yarl.URL(uri))
    async with aiohttp.ClientSession() as session:
        async with session.get(str(yarl.URL(uri)), timeout=timeout) as json_request:
            return json.loads(await json_request.text())

# -----------------------------------------------------
# ---------- Spreadsheet processor functions ----------
# -----------------------------------------------------

async def dt_0(data, message):
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

    # header = wb.add_format({'bold': True, 'font_size': 24, 'center_across': True, 'border': 2})
    fill = wb.add_format({'fg_color': '#000000'})
    misc_cell = wb.add_format({'font_name': 'Arial', 'font_size': 10, 'center_across': True, 'bg_color': '#999999'})

    ws.write(f'A1:F{len(data) + 4}', ' ', fill)

    ws.write('B2', 'Character Name', misc_cell)
    ws.write('C2', 'Invite Status', misc_cell)
    ws.write('D2', 'Equipped', misc_cell)
    ws.write('E2', 'Neck', misc_cell)

    class_cell = wb.add_format(cell_format['Priest'])
    invite_cell = wb.add_format(cell_format['Signed Up'])

    ws.write('B3', 'TASIN FILIP', class_cell)
    ws.write('C3', 'CONFIRM KO KUĆA', invite_cell)
    ws.write('D3', '370-380', misc_cell)
    ws.write('E3', 'Oko 31', misc_cell)

    # Populate the rows
    info = await message.author.send('Getting character data...')

    embed = discord.Embed(
        title=f'Spreadsheet preview',
        description='Quick run-down of the roster below',
        color=0xfaa61a
    )
    embed_conf = ''
    embed_tent = ''
    conf_count = 0
    tent_count = 0

    for i in range(1, len(data) + 1):

        i = str(i)
        # Assign variables for increased readability
        inv_stat = invite_status[data[i]['stat']]
        class_name = data[i]['cls']
        char_name = data[i]['name']

        try:
            char_name, realm = char_name.split('-')
        except:
            realm = 'kazzak'

        await info.edit(
            content=f'`[{int(i)}/{len(data)}]` Getting info for {char_name}-{realm.title()}\n'
            f'`[{"█" * (int(int(i) / len(data) * 25)) + "-" * (25 - int(int(i) / len(data) * 25))}]`'
        )

        print(int(int(i) / len(data) * 10))

        character_data = await get_json(
            f'https://eu.api.blizzard.com/wow/character/{realm}/{char_name}?fields=items&locale=en_GB&access_token={bnet_token}'
        )

        # Cell value set to character's name, cell style set to character's class (class color cell BG for now only)
        class_cell = wb.add_format(cell_format[class_name])
        invite_cell = wb.add_format(cell_format[inv_stat])

        char_eq = character_data['items']['averageItemLevelEquipped']
        char_hoa = character_data['items']['neck']['azeriteItem']['azeriteLevel']
        i = int(i)

        ws.write(f'B{str(i + 3)}', char_name, class_cell)
        ws.write(f'C{str(i + 3)}', inv_stat, invite_cell)
        ws.write(f'D{str(i + 3)}', char_eq, misc_cell)
        ws.write(f'E{str(i + 3)}', char_hoa, misc_cell)

        if inv_stat == "Signed Up" or inv_stat == "Confirmed":
            conf_count +=1
            embed_conf += f'{char_name} [{char_eq}/{char_hoa}]\n'
        elif inv_stat == "Tentative":
            tent_count += 1
            embed_tent += f'{char_name} [{char_eq}]/[{char_hoa}]\n'

    embed.add_field(
        name=f'```Confirmed [{conf_count}]```',
        value=f'```{embed_conf}```',
        inline=False
    )

    embed.add_field(
        name=f'```Tentative [{tent_count}]```',
        value=f'```{embed_tent if tent_count != 0 else "No tentative players"}```',
        inline=False
    )
    await info.edit(content=info.content + f'\nFinished gathering data, {message.author.mention}', embed=embed)
    # Close the sheet
    wb.close()

    return wb_name


async def dt_1(data, message):
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

    ws.write(f'A1:C{len(all_online) + 3}', ' ', fill)
    ws.write('B2', 'Online Characters', header)
    # Populate the rows
    for i in range(len(all_online)):
        char_name = all_online[i]
        ws.write(f'B{i + 3}', char_name, body)

    wb.close()

    return wb_name


async def dt_2(data, message):
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

    ws.write(f'A1:E{len(data) + 3}', ' ', fill)
    ws.write('B2', 'Character Name', header)
    ws.write('C2', 'Member Note', header)
    ws.write('D2', 'Officer Note', header)
    # Populate the rows
    for i in range(1, len(data)):
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
        ws.write(f'B{str(i + 2)}', char_name, body)
        ws.write(f'C{str(i + 2)}', member_note, body)
        ws.write(f'D{str(i + 2)}', officer_note, body)

    # Close the sheet
    wb.close()

    return wb_name


async def dt_3(data, message):
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

    ws.write(f'A1:C{len(all_in_raid) + 3}', ' ', fill)
    ws.write('B2', 'Characters in raid', header)
    # Populate the rows
    for i in range(len(all_in_raid)):
        char_name = all_in_raid[i]

        ws.write(f'A{i + 3}', ' ', fill)
        ws.write(f'B{i + 3}', char_name, body)
        ws.write(f'C{i + 3}', ' ', fill)

    wb.close()
    return wb_name


# -----------------------------------------------------
# -----------------------------------------------------
class EventExport(discord.ext.commands.Bot):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.version = '1.1.20'
        self.owner = None
        # Spreadsheet processors (data, message)
        self.ssp = [
            dt_0,
            dt_1,
            dt_2,
            dt_3,
        ]

    async def on_ready(self):
        await self.change_presence(activity=discord.Game(name='Files and Base64 :3'))
        self.owner = self.get_user(217294231125884929)
        await self.owner.send(f'`[{datetime.datetime.now():%H:%M:%S}]` Connected to Discord!')

        print('---------------------')
        print(f'Logged in as: {self.user.name} - {self.user.id}')
        print(f'Version: {self.version}')
        print('---------------------')

    async def on_message(self, message):
        if message.content.startswith('!'):
            await self.process_commands(message)
            return

        # Escape bot message
        if message.author == bot.user:
            return

        # Inform me (Bot coder) of people using the bot
        if message.author != self.owner:
            file = None
            if message.attachments:
                att = message.attachments[0]
                await att.save(att.filename)
                file = discord.File(att.filename)

            await self.owner.send(
                f'`[{datetime.datetime.now():%H:%M:%S}]` {message.author}: ```{message.content if message.content != "" else "N/A"}```', file=file)

        # Read data from file attached (happens when sting has more than 2000 chars)
        if message.attachments:
            file = message.attachments[0]  # Expecting only one attachment
            async with message.channel.typing():
                print(f'{message.author} requested a spreadsheet parse with import file: {file.filename}')
                # Prepare string for parsing
                await file.save('tempfile.txt')
                with open('tempfile.txt', 'rb') as openfile:
                    importstring = openfile.readlines()
                    try:
                        read_data = [line.decode('utf_8') for line in importstring][0]
                    except UnicodeDecodeError:
                        await message.author.send("Error:```UnicodeDecodeError: 'utf-8' codec can't decode the file provided```")
                        return
        # Read message data (has less than 2000 characters)
        else:
            print(f'{message.author} requested a spreadsheet parse with import: {message.content}')
            read_data = message.content

        # Decode the string
        async with message.channel.typing():
            try:
                decoded_data = b64.b64decode(read_data)
            except ValueError:
                await message.author.send("Error:```UnicodeDecodeError: 'utf-8' codec can't decode the string provided```")
                return

            # Convert to JSON
            data = json.loads(decoded_data)
            print(data)

            # Extracts "stringType" from dict into a value
            if data['stringType']:
                data_type = data['stringType']
                del data['stringType']
            else:
                await message.author.send('Invalid JSON received :S')
                return

            # Makes the spreadsheet
            try:
                spreadsheet = await self.ssp[int(data_type)](data, message)
                # Send the file back
                await message.author.send(file=discord.File(spreadsheet))
            except Exception as e:
                print(e)
                traceback.print_exc()

    async def on_command_error(self, ctx, error):
        print(ctx, error)


# Bot wants ManageMessages, ReadChannelHistory
def get_prefix(bot, message):
    prefixes = ['!']
    if not message.guild:
        return '!'
    return commands.when_mentioned_or(*prefixes)(bot, message)


initial_extensions = [
    'cogs.b64'
]

bot = EventExport(command_prefix=get_prefix, description='EventExport')

if __name__ == '__main__':
    for extension in initial_extensions:
        try:
            print(extension)
            bot.load_extension(extension)
        except Exception as e:
            print(f'Failed to load extension {extension}.', file=sys.stderr)
            traceback.print_exc()

bot.run(os.environ["BOT_TOKEN"], bot=True, reconnect=True)
