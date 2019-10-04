import unittest
import slovnik

class TestDictionary(unittest.TestCase):

    def setUp(self):
        self.tested_dict = slovnik.Dictionary()

    def tearDown(self):
        pass
    
    def test_show(self): 
        self.assertTrue(self.tested_dict.show(prelanguage=1))
        self.assertTrue(self.tested_dict.show(prelanguage=2))

    def test_delete(self):
    	self.assertTrue(self.tested_dict.delete("der Hund"))

    def test_insert_a_word(self):
        self.assertTrue(self.tested_dict.insert_a_word("das", "Lied", "píseň"))
        self.assertFalse(self.tested_dict.insert_a_word("*", "píseň"))
        self.assertFalse(self.tested_dict.insert_a_word("", "píseň"))
        self.assertFalse(self.tested_dict.insert_a_word("", ""))
        self.assertFalse(self.tested_dict.insert_a_word("", "*"))
        self.assertFalse(self.tested_dict.insert_a_word("1", "*"))
        self.assertFalse(self.tested_dict.insert_a_word("e", "-1"))
        self.assertTrue(self.tested_dict.insert_a_word("die", "Hilfe", "pomoc"))

if __name__ == '__main__':
    unittest.main()