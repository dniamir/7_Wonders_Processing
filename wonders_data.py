import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import openpyxl as xl 

class wonder_game():
    """Class containing all data from a single 7 Wonders game"""
    def __init__(self, xlsx_filepath, sheet_name):

        # Format raw data
        df = pd.read_excel(xlsx_filepath, sheet_name=sheet_name)
        df.set_index(list(df)[0], inplace=True)

        # Save meta data
        self.filepath = xlsx_filepath
        self.sheet_name = sheet_name
        self.players = list(df)
        self.n_players = len(self.players)
        self.recorded_civ = False
        self.victor_civ = None
        self.victor_side = None
        self.victor_civside = None

        # Ensure total is correct
        for player in self.players:
            a = df.loc['Red', player]
            b = df.loc['Coins', player]
            c = df.loc['Wonders', player]
            d = df.loc['Blue', player]
            e = df.loc['Yellow', player]
            f = df.loc['Purple', player]
            g = df.loc['Green', player]
            df.loc['Total', player] = a + b + c + d + e + f + g

        self.raw_data = df        
        self.victor = self.raw_data.iloc[1:, :].idxmax(axis=1)['Total']

        # Check if civs were recorded
        if isinstance(self.raw_data.loc['Civilization', :].values[0], str):
            self.recorded_civ = True

        # Save civilization data
        if self.recorded_civ:
            civs, sides, civ_sides = self.get_civs()
            self.raw_data = self.raw_data.iloc[1:, :]
            self.raw_data.loc['Civ'] = civs
            self.raw_data.loc['Side'] = sides
            self.raw_data.loc['CivSide'] = civ_sides
            self.victor_civ = self.raw_data.loc['Civ', self.victor]
            self.victor_side = self.raw_data.loc['Side', self.victor]
            self.victor_civside = self.raw_data.loc['CivSide', self.victor]
        else:
            self.raw_data = self.raw_data.iloc[1:, :]
            self.raw_data.loc['Civ'] = None
            self.raw_data.loc['Side'] = None
            self.raw_data.loc['CivSide'] = None

    def get_civs(self):
        """"Return civilization data from the game"""
        civs = self.raw_data.loc['Civilization', :].values

        sides = [civ[-1] for civ in civs]
        civ_sides = [civ.replace('-', '').replace(' ', '') for civ in civs]
        civs = [civ.split('-')[0].replace(' ', '') for civ in civs]

        return civs, sides, civ_sides
    
class all_wonder_games():
    """Class containing data from multiple 7 Wonders games"""
    def __init__(self, xlsx_filepath):
        sheet_names = xl.load_workbook(xlsx_filepath).sheetnames
        self.all_games = {}
        self.all_players = {}

        # Import data from all games
        for sheet_name in sheet_names:
            game = wonder_game(xlsx_filepath, sheet_name=sheet_name)

            # Get cumulative player data
            for player in game.players:
                
                player_df = game.raw_data[[player]].T
                if player not in list(self.all_players):
                    self.all_players[player] = player_df
                else:
                    self.all_players[player] = pd.concat([self.all_players[player], player_df])
            
            self.all_games[sheet_name] = game
        