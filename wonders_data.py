import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import openpyxl as xl 

class game():
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

        # Get high level stats
        self.civs, self.sides, self.civ_sides = self.get_civs()

    def get_civs(self):
        """"Return civilization data from the game"""
        civs = self.raw_data.loc['Civilization', :].values

        # Ensure that the civilizations were actually entered
        if not isinstance(civs[0], str):
            return civs, civs, civs

        sides = [civ[-1] for civ in civs]
        civ_sides = [civ.replace('-', '').replace(' ', '') for civ in civs]
        civs = [civ.split('-')[0].replace(' ', '') for civ in civs]

        return civs, sides, civ_sides
    
class cumulative_games():
    """Class containing data from multiple 7 Wonders games"""
    def __init__(self, xlsx_filepath):
        sheet_names = xl.load_workbook(xlsx_filepath).sheetnames
        self.all_games = {}

        # Import data from all games
        for sheet_name in sheet_names:
            self.all_games[sheet_name] = game(xlsx_filepath, sheet_name=sheet_name)

            # Get cumulative player data
        