' Test GetUrlFromCommentPermLink

Dim x
Set x = New Utility

AssertEqual x.GetUrlFromCommentPermLink("re-tvb-re-justyy-re-tvb-45qr3w-20171011t144205534z"), "https://steemit.com/@tvb/45qr3w", ""

AssertEqual x.GetUrlFromCommentPermLink("re-justyy-daily-quality-cn-posts-selected-and-rewarded-promo-cn-20180520t153728557z"), "https://steemit.com/@justyy/daily-quality-cn-posts-selected-and-rewarded-promo-cn", ""

Set x = Nothing