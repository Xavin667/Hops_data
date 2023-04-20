"""Test file for the main.py"""
import unittest
from unittest import mock
from Hops_data.main import *


class TestHops(unittest.TestCase):
    @mock.patch('main.input', create=True)
    def test_get_hop_name(self, mock_input):
        mock_input.return_value = 'Columbus'
        self.assertEqual(get_hop_name(), 'Columbus')
        mock_input.return_value = 'Waiiti'
        self.assertEqual(get_hop_name(), 'Waiiti')

    def test_get_sheet_name(self):
        self.assertEqual(get_sheet_name('Nowa Zelandia'), nz)
        self.assertEqual(get_sheet_name('USA'), usa)
        self.assertEqual(get_sheet_name('Australia'), au)
        self.assertEqual(get_sheet_name('Polska'), pl)
        self.assertEqual(get_sheet_name('Czechy'), None)

    def test_print_aromas(self):
        self.assertEqual(print_aromas(nz, nz.cell(row=2, column=1)), (['Owoce tropikalne',
            'Igły sosnowe', 'Nuty kwiatowe'], ['Limonka', 'Igły sosnowe', 'Ananas', 'Jagody']))
        self.assertEqual(print_aromas(nz, nz.cell(row=23, column=1)), (['Słodkie owoce tropikalne',
            'Agrest', 'Białe winogrona'], ['Agrest', 'Białe winogrona (wino)', 'Marakuja',
            'Limonka', 'Igły sosnowe']))
        self.assertEqual(print_aromas(au, au.cell(row=10, column=1)), (['Owoce tropikalne', 'Jagody', 'Owoce pestkowe'],
            ['Malina', 'Czerwona porzeczka', 'Białe winogrona (wino)', 'Melon cukrowy (kantalupa)']))
        self.assertEqual(print_aromas(usa, usa.cell(row=61, column=1)), (['Owoce tropikalne, cytrusowe', 'Żywiczny',
            'Korzenny', 'Nuty kwiatowe'], ['Jagoda', 'Mandarynka', 'Papaja', 'Róża', 'Guma balonowa', 'Trawa']))
        self.assertEqual(print_aromas(pl, pl.cell(row=2, column=1)), (['Kwiatowy', 'Ziołowe, trawiaste', 'Korzenne'],
            ['Bergamotka', 'Cynamon', 'Magnolia', 'Lawenda', 'Herbata z cytryną']))


if __name__ == '__main__':
    unittest.main()