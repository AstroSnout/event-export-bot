from discord.ext import commands
import base64 as b64
import ast
import json


class B64:
    def __init__(self, bot):
        self.bot = bot

    @staticmethod
    async def _pretty_json(text):
        # Replace lua's 'true' with py's 'True'
        # And get a dict from string representation
        text = ast.literal_eval(text.replace('true', 'True'))
        return json.dumps(text, indent=4, sort_keys=True)

    @commands.command(name='encode', aliases=['e', 'enc'], help='Encodes string to Base64')
    async def encode(self, ctx, *, string):
        await ctx.send(f'```{b64.b64encode(string.encode()).decode()}```')

    @commands.command(name='decode', aliases=['d', 'dec'], help='Decodes string from Base64')
    async def decode_base64(self, ctx, *, string):
        print(string)
        missing_padding = len(string) % 4
        if missing_padding != 0:
            string += b'=' * (4 - missing_padding)
        decoded = b64.b64decode(string).decode()
        try:
            pretty_decoded = await B64._pretty_json(decoded)
        except ValueError:
            pretty_decoded = decoded

        if len(pretty_decoded) > 2000:
            for i in range(len(pretty_decoded)//1994+1):
                await ctx.send(f'```{pretty_decoded[1994*i:1994*(i+1)-1]}```')
        else:
            await ctx.send(f'```{pretty_decoded}```')


def setup(bot):
    bot.add_cog(B64(bot))
