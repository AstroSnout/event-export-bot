import discord
from discord.ext import commands

import sys
import traceback
import base64 as b64
import json
import datetime
import ast
from ss_maker import Spreadsheets

def get_prefix(bot, message):
    prefixes = ['!']
    if not message.guild:
        return '!'
    return commands.when_mentioned_or(*prefixes)(bot, message)


async def get_owner(bot):
    return bot.get_user(217294231125884929)

# Below cogs represents our folder our cogs are in. Following is the file name. So 'meme.py' in cogs, would be cogs.meme
# Think of it like a dot path import
initial_extensions = []

bot = commands.Bot(command_prefix=get_prefix, description='EventExport')
ss_maker = Spreadsheets()

# Here we load our extensions(cogs) listed above in [initial_extensions].
if __name__ == '__main__':
    for extension in initial_extensions:
        try:
            bot.load_extension(extension)
        except Exception as e:
            print(f'Failed to load extension {extension}.', file=sys.stderr)
            traceback.print_exc()


async def encode_base64(text):
    return b64.b64encode(text.encode()).decode()


async def decode_base64(data):
    missing_padding = len(data) % 4
    if missing_padding != 0:
        data += b'='* (4 - missing_padding)
    return b64.b64decode(data).decode()


async def pretty_json(text):
    # Replace lua's 'true' with py's 'True'
    # And get a dict from string representation
    text = ast.literal_eval(text.replace('true', 'True'))
    return json.dumps(text, indent=4, sort_keys=True)

@bot.event
async def on_ready():
    await bot.change_presence(activity=discord.Game(name='Files and Base64 :3'))
    me = await get_owner(bot)
    await me.send('I am ready to accept inputs')


@bot.event
async def on_message(message):
    me = await get_owner(bot)

    if message.author == bot.user:
        return

    if message.content.startswith('encode '):
        encoded = await encode_base64(message.content[7:])
        await message.author.send('```' + encoded + '```')
        return

    if message.content.startswith('decode '):
        decoded = await decode_base64(message.content[7:])
        decoded = await pretty_json(decoded)
        await message.author.send('```'+decoded+'```')
        return

    if message.author != me:
        await me.send(f'```[{datetime.datetime.now().strftime("%H:%M:%S")}]{message.author}: {message.content}```')

    if message.attachments:
        file = message.attachments[0]  # Expecting only one attachment
        async with message.channel.typing():
            print(f'{message.author} requested a spreadsheet parse with import file: {file.filename}')
            # Prepare string for parsing
            await file.save('tempfile.txt')
            with open('tempfile.txt', 'rb') as openfile:
                importstring = openfile.readlines()
                read_data = [line.decode('utf_8') for line in importstring][0]
    else:
        print(f'{message.author} requested a spreadsheet parse with import: {message.content}')
        read_data = message.content

    async with message.channel.typing():
        try:
            decoded_data = b64.b64decode(read_data)
        except ValueError:
            await message.author.send('Invalid input string')
            return

        data = json.loads(decoded_data)
        print(data)
        if data['stringType']:
            data_type = data['stringType']
            del data['stringType']
        else:
            await message.author.send('Invalid JSON received :S')
            return

        spreadsheet = ss_maker.make_spreadsheet(data_type, data)
        # Send the file back
        await message.author.send(file=discord.File(spreadsheet))


bot.run('NDg1MDkwNTM4ODE1NzUwMTc1.DmrfgQ.nrPA_zao8CQUyjdbdSKhriuxvTg', bot=True, reconnect=True)
