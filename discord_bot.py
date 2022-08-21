import discord


class MyClient(discord.Client):
    async def on_ready(self):
        print('Logged on as {0}!'.format(self.user))
        channel = self.get_channel(940195025051090995)
        # await channel.purge(limit = 2)
        await self.send_tt()
        await self.close()
    
    async def on_message(self, message):
        if message.content.startswith('$hello'):
            await message.channel.send('Yeah?')
        elif message.content == '$tt':
            channel = self.get_channel(940195025051090995)
            await channel.send(file=discord.File('timetable.png'))
        else:
            print('Message from {0.author}: {0.content}'.format(message))
    
    async def send_tt(self):
        channel = self.get_channel(940195025051090995)
        await channel.send(file=discord.File('timetable.png'))
