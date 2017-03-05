import movingaveragecalc
import configparser
import unittest
import gspread
from unittest.mock import Mock, patch


class MovingAverageCalcTest(unittest.TestCase):
    def setUp(self):
        configpath = 'config.txt'
        configParser = configparser.RawConfigParser()
        configParser.read(configpath)
        self.credentials_json = str(configParser.get('arguments', 'credentials_json'))
        self.calc_window = int(configParser.get('arguments', 'calc_window'))
        self.ssheet_id = configParser.get('arguments', 'ssheet_Id')
        self.scope = configParser.get('arguments', 'scope')

    def test_get_data_no_data(self):
        creds = self.create_patch('oauth2client.service_account.ServiceAccountCredentials.from_json_keyfile_name')
        creds.return_value = ''
        gc = self.create_patch('gspread.authorize')
        gc.return_value.open_by_key.return_value.get_worksheet.return_value.get_all_values.return_value = []
        with patch('sys.stdout', new=Mock()) as fake_out:
            self.assertRaises(ValueError, movingaveragecalc.get_data, '', self.ssheet_id, self.scope)

    def test_get_data_wrongId(self):
        creds = self.create_patch('oauth2client.service_account.ServiceAccountCredentials.from_json_keyfile_name')
        creds.return_value = ''
        gc = self.create_patch('gspread.authorize')
        gc.return_value.open_by_key.side_effect = gspread.exceptions.SpreadsheetNotFound
        self.assertRaises(gspread.exceptions.SpreadsheetNotFound, movingaveragecalc.get_data, self.credentials_json, 'WrongId', self.scope)

    def test_get_data(self):
        creds = self.create_patch('oauth2client.service_account.ServiceAccountCredentials.from_json_keyfile_name')
        creds.return_value = ''
        gc = self.create_patch('gspread.authorize')
        gc.return_value.open_by_key.return_value.\
            get_worksheet.return_value.get_all_values.return_value = [['Visitors', 'Date']]
        with patch('sys.stdout', new=Mock()) as fake_out:
            res_ws, res_headers, res_data = movingaveragecalc.get_data(self.credentials_json, self.ssheet_id, self.scope)
        self.assertEqual(res_headers, ['Visitors', 'Date'])
        self.assertEqual(res_data, [])

    def create_patch(self, *args):
        get_patcher = patch if len(args) == 1 else patch.object
        patcher = get_patcher(*args)
        res = patcher.start()
        self.addCleanup(patcher.stop)
        return res

    def test_cast_int(self):
        self.assertEqual(movingaveragecalc.cast_int(24, 0), 24)
        self.assertEqual(movingaveragecalc.cast_int('24.4', 0), 0)
        self.assertEqual(movingaveragecalc.cast_int('aaa', 0), 0)
        self.assertEqual(movingaveragecalc.cast_int('0', 1), 0)

    def test_moving_average(self):
        res = movingaveragecalc.moving_average([1, 1, 1, 2, 2, 2, 3, 3, 3], 3)
        self.assertListEqual(res, [1, 1, 2, 2, 2, 3, 3])
        res = movingaveragecalc.moving_average([1, 2, 3, 4, 5, 6, 7], 3)
        self.assertListEqual(res, [2, 3, 4, 5, 6])

    def test_get_col_name(self):
        self.assertEqual('A', movingaveragecalc.get_col_name(1))
        self.assertEqual('CZ', movingaveragecalc.get_col_name(4*26))
        self.assertEqual('ALL', movingaveragecalc.get_col_name(1000))


if __name__ == '__main__':
    suite = unittest.TestLoader().loadTestsFromTestCase(MovingAverageCalcTest)
    unittest.TextTestRunner(verbosity=2).run(suite)
