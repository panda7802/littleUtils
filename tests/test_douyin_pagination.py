import unittest
from unittest.mock import patch

import douyin_account_videos as douyin


class FakeResponse:
    url = "https://www.douyin.com/aweme/v1/web/aweme/post/?max_cursor=1"
    status = 200

    def __init__(self, items, has_more):
        self.items = items
        self.has_more = has_more

    def json(self):
        return {"aweme_list": self.items, "has_more": self.has_more, "status_code": 0}


class FakePage:
    def __init__(self):
        self.listener = None

    def on(self, event, callback):
        self.listener = callback

    def remove_listener(self, event, callback):
        self.listener = None

    def goto(self, *args, **kwargs):
        return None

    def wait_for_timeout(self, milliseconds):
        return None


class PaginationTest(unittest.TestCase):
    def test_collects_all_434_posts_before_confirmed_end(self):
        page = FakePage()
        pages = [
            [{"aweme_id": str(number), "desc": f"作品 {number}", "create_time": 1}
             for number in range(start, min(start + 20, 434))]
            for start in range(0, 434, 20)
        ]
        current = 0

        def advance(fake_page):
            nonlocal current
            if current < len(pages):
                fake_page.listener(FakeResponse(pages[current], current < len(pages) - 1))
                current += 1

        with patch.object(douyin, "advance_page", side_effect=advance):
            rows, complete = douyin.collect_posts(page, "https://www.douyin.com/user/example", 10_000, 60)
        self.assertEqual(len(rows), 434)
        self.assertTrue(complete)


if __name__ == "__main__":
    unittest.main()
