import sc2reader
from openpyxl import Workbook
from openpyxl.worksheet.dimensions import ColumnDimension, DimensionHolder
from openpyxl.utils import get_column_letter

class Game():
    def __init__(self, replay):
        self.map = replay.map
        self.file_name = replay.filename
        self.type = replay.type
        self.datetime = replay.date
        self.category = replay.category
        self.length = replay.length.seconds
        self.build = replay.build
        self.release_string = replay.release_string
        self.teams = [team for team in replay.teams]
        self.matchup = "v".join(sorted(team.lineup for team in self.teams))
        self.players = sum((team.players for team in self.teams), [])

def fixColumnWidth(ws):
    dim_holder = DimensionHolder(worksheet=ws)

    for col in range(ws.min_column, ws.max_column + 1):
        dim_holder[get_column_letter(col)] = ColumnDimension(ws, min=col, max=col, width=20)

    ws.column_dimensions = dim_holder

replays = sc2reader.load_replays('replays/', load_level=2, load_map=True)

wb = Workbook()
overviewWS = wb.active
overviewWS.title = "Overview"
overviewWS.append(['Name', 'Win', 'Loss', 'Terran', 'Zerg', 'Protoss'])

gamesWS = wb.create_sheet("Games")
gamesWS.append(['Date', 'Map', 'Type', 'Length', 'Team 1', 'Team 2'])

playersWS = wb.create_sheet("Players")
playersWS.append(['Date', 'Map', 'Name', 'Race', 'Result'])

all_players = []

for replay in replays:    
    game = Game(replay)

    if game.type == '1v1':
        continue

    print(f'{game.datetime} - {game.map.name} {game.type}, Game length: {round(game.length / 60)}:{game.length % 60}')

    game_result = []

    for team in game.teams:
        players = ", ".join([f'{player.name} ({player.play_race})' for player in team.players])
        game_result.append(f'{team.number} - {team.result} - {players}')

        for player in team.players:
            playersWS.append([game.datetime, game.map.name, player.name, player.play_race, team.result])
            if (player.name not in all_players):
                all_players.append(player.name)

    gamesWS.append([game.datetime, game.map.name, game.type, f'{round(game.length / 60)}:{game.length % 60}'] + game_result)

for index, player in enumerate(all_players):
    overviewWS.append([player, 
                       f'=COUNTIFS(Players!$C:$C,$A{index+2},Players!$E:$E,"Win")', 
                       f'=COUNTIFS(Players!$C:$C,$A{index+2},Players!$E:$E,"Loss")',
                       f'=COUNTIFS(Players!$C:$C,$A{index+2},Players!$D:$D,"Terran")',
                       f'=COUNTIFS(Players!$C:$C,$A{index+2},Players!$D:$D,"Zerg")',
                       f'=COUNTIFS(Players!$C:$C,$A{index+2},Players!$D:$D,"Protoss")',
                    ])

fixColumnWidth(overviewWS)
fixColumnWidth(gamesWS)
fixColumnWidth(playersWS)

overviewWS.freeze_panes = overviewWS['A2']
gamesWS.freeze_panes = gamesWS['A2']
playersWS.freeze_panes = playersWS['A2']

# Save the file
wb.save("sc2.xlsx")