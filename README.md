GAS(Google Apps Script)を使用して公開されているGoogleカレンダーのカレンダーID、DiscordのWeb hookのURLを指定することで以下のようなメッセージをDiscordに送信できます。

![メッセージの例](https://github.com/murahito130/google_calendar_freebusy_discord_webhook/blob/main/example.png)


このコードをGASのトリガーを使用して数時間(4時間以内を推奨)毎に実行することで同じメッセージを編集する形で1つのメッセージを更新できます。
通知を煩わしく思う方にオススメです。

# 注意
制作者の生活サイクルを想定してプログラムされています。
具体的には、土日祝日は予定が空いており、12時から18時がお昼、20時から25時を夜としております。
生活サイクルが異なる方や、変更を加えたい方はご自由にカスタマイズしてください。
