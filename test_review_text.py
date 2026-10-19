import unittest
from review_text import segments


class ReviewTextTests(unittest.TestCase):
    def test_word_diff_preserves_complete_arabic_and_untrusted_text(self):
        a='ME-01 توريد مضخة 10 <script>\nمرحبا';b='ME-01 توريد مضخة 20 <script>\nمرحبا'
        sa,sb=segments(a,b)
        self.assertEqual(''.join(p['text'] for p in sa),a)
        self.assertEqual(''.join(p['text'] for p in sb),b)
        self.assertEqual([p['text'] for p in sa if p['changed']],['10'])
        self.assertEqual([p['text'] for p in sb if p['changed']],['20'])
    def test_very_long_text_uses_bounded_fallback(self):
        a='بند '*10000;b=a+'إضافة'
        sa,sb=segments(a,b)
        self.assertEqual(len(sa),1);self.assertEqual(sa[0]['text'],a);self.assertEqual(sb[0]['text'],b)
