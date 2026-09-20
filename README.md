# 抖音、B站、小红书公开作品采集工具

输入抖音号，采集该公开账号的作品页链接、标题、完整文案、发布时间，以及页面响应中可用时的播放/封面地址。脚本用真实 Chrome 打开抖音并监听网页自身的分页请求，因此不需要自行实现易失效的 `a_bogus` / `X-Bogus` 签名。结果直接输出为带中文表头的 Excel 文件，文件名使用账号昵称。

## 安装

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
playwright install chromium
```

脚本默认使用本机已安装的 Chrome（`channel="chrome"`）；如果机器没有 Chrome，可删除代码中的 `channel="chrome"`，使用 Playwright 安装的 Chromium。

## 示例：沈建平（抖音号 837672563）

```powershell
python douyin_account_videos.py 837672563
```

第一次运行会打开 Chrome。若抖音要求登录或验证码，请在浏览器完成；登录状态保存在本地 `.douyin-browser` 中。脚本会自动搜索抖音号、进入匹配的主页、滚动至末尾并输出，例如：

默认等待手动登录 120 秒，可用 `--login-wait 300` 延长等待时间。

- `output/沈建平.xlsx`

如果自动搜索无法精确定位，在抖音网页打开对应主页并复制 URL：

```powershell
python douyin_account_videos.py 837672563 --profile-url "https://www.douyin.com/user/用户的sec_uid"
```

调试时可先限制数量：

```powershell
python douyin_account_videos.py 837672563 --max-posts 20
```

## Excel 列说明

- `作品链接`：稳定的作品页面地址，建议业务中保存此列。
- `作品类型`：视频或图文。
- `标题`：优先读取页面返回的标题字段；没有独立标题时取文案首行。
- `内容`：作品完整文案。
- `播放地址（临时）`：CDN 临时播放地址，可能过期或要求 Cookie/Referer，不应作为永久链接。
- `图片地址`：图文作品的图片地址。

## 限制与合规

官方 OpenAPI 的视频列表接口仅适用于账号本人 OAuth 授权后的数据，不能仅凭第三方抖音号调用。本工具只读取登录用户在网页上可见的公开内容；请控制频率，并遵守平台规则、个人信息与著作权要求。页面结构或接口变化时，脚本可能需要相应更新。

## B站公开视频导出

传入 UP 主 UID 或空间链接：

```powershell
python bilibili_account_videos.py 946974
python bilibili_account_videos.py "https://space.bilibili.com/946974"
```

结果保存在 `output/bilibili/UP主昵称.xlsx`。表格包含 BV/AV 号、标题、简介、发布时间、时长、公开视频链接、封面以及页面返回的播放、弹幕、评论和收藏数。登录状态保存在 `.bilibili-browser`。

## 小红书公开笔记导出

推荐直接传入用户主页 URL：

```powershell
python xiaohongshu_account_notes.py "https://www.xiaohongshu.com/user/profile/用户ID"
```

也可以使用昵称或小红书号搜索，但重名时建议改传主页 URL：

```powershell
python xiaohongshu_account_notes.py "昵称或小红书号"
```

结果保存在 `output/xiaohongshu/作者昵称.xlsx`。默认逐条打开笔记补齐正文；只需要主页列表、希望加快速度时可使用：

```powershell
python xiaohongshu_account_notes.py "用户主页URL" --skip-details
```

主页列表采集完成后会立即写入一次 Excel 检查点，逐条补正文时每 10 条更新一次。中途关闭浏览器时，脚本会停止补正文并保留当前已经获取的数据。

登录状态保存在 `.xiaohongshu-browser`。首次运行若出现登录或验证码，请在打开的 Chrome 中手动完成，脚本不会绕过平台验证。
# littleUtils
