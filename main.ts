// スプレッドシート 時間外勤務表の管理
const ssID = '1EhXDA7gHDe7hNZqbjDhERZabgYCrSsvIXaL-ohRr2Eg'
const ss: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ssID)
const sheet: GoogleAppsScript.Spreadsheet.Sheet | null = ss.getSheetByName('今月')
const scriptProperties: GoogleAppsScript.Properties.Properties = PropertiesService.getScriptProperties()

// トリガーで毎日AM1:00に実行する
function main() {
	console.log('This is Main function !')

	// 1日かどうかチェックする．
	if (isCheckDate()) {
		// シートをコピー
		copySheet()
		const fileName = createFileName()
		const pdfBlob = createPdfBlob(ss, fileName)
		const pdfFile = createPdfFile(pdfBlob)
		const fileUrl = getFileUrl(pdfFile)

		// メールに送信
		sendEmail(pdfBlob, fileUrl)

		// LINEに送信
		const lineNotifyMessage = `\n\n今月の時間外勤務表\n\n${fileUrl}`
		sendLineNotify(lineNotifyMessage)
	} else {
		console.log('今日は1日じゃないので関数実行しませんでした．')
	}
}

function sendLineNotify(message: string = 'テスト通知です') {
	const LINE_NOTIFY_API_TOKEN = scriptProperties.getProperty('LINE_NOTIFY_API_TOKEN')
	const LINE_NOTIFY_API_URL = 'https://notify-api.line.me/api/notify'

	const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
		method: 'post',
		headers: {
			'Content-Type': 'application/x-www-form-urlencoded',
			Authorization: `Bearer ${LINE_NOTIFY_API_TOKEN}`,
		},
		payload: { message },
	}

	// LINEに通知
	UrlFetchApp.fetch(LINE_NOTIFY_API_URL, options)
}

// POSTリクエストが来たときの処理 ---------------------------------------------------
function doPost(e: GoogleAppsScript.Events.DoPost) {
	// デバッグlog
	console.log('e:  ', e)

	// ReactAppからPOSTされたデータを取得する
	const data: OvertimePostData = JSON.parse(e.postData.contents).data
	let { radiologist, modality, date, start, end, description } = data

	// タイムスタンプ生成
	const created_at = new Date()

	//! スプレッドシートにフォームデータを書き込んで、lineNotify関数を実行する．
	// スプレッドシートに追加する項目↓
	// [uuid, created_at, created_at, date, radiologist, modality, start, end, description]
	const uuid = Utilities.getUuid()

	// date,start,endはstring型なのでDate型に変換する．
	const parsed_date = new Date(date)
	const parsed_start = toDateFromStrTime(start)
	const parsed_end = toDateFromStrTime(end)
	const diff_time = ((parsed_end.getTime() - parsed_start.getTime()) / 1000 / 60).toFixed(1) // 分
	const diff_time_2 = ((parsed_end.getTime() - parsed_start.getTime()) / 1000 / 60 / 60).toFixed(1) // 時間

	// 10列目に残業時間（分）のヘッダーが無ければ 入力する．
	const total_time_cell = sheet?.getRange(1, 10)
	if (total_time_cell?.getValue() === '') {
		total_time_cell.setValue('残業時間（分）')
	}
	// 11列目に残業時間（時間）のヘッダーが無ければ 入力する．
	const total_time_cell_2 = sheet?.getRange(1, 11)
	if (total_time_cell_2?.getValue() === '') {
		total_time_cell_2.setValue('残業時間（分）')
	}

	// spreadsheetに追加する．
	sheet?.appendRow([
		uuid,
		created_at,
		created_at,
		radiologist,
		modality,
		parsed_date,
		start,
		end,
		description,
		diff_time,
		diff_time_2,
	])
	// sheet?.appendRow([uuid, created_at, created_at, radiologist, modality, date, start, end, description, diff_time])

	// スプレッドシートに書き込まれるまで少し待機
	Utilities.sleep(5 * 1000)

	// const testDataMessage = `これはテストデータです．\n\n${e.postData.contents}`
	const postMessage = `時間外勤務を登録しました．\n\n${date}\n${radiologist}\n${modality}\n${start} 〜 ${end}\n${description}`
	// LINEへ通知
	sendLineNotify(postMessage)
	// lineNotifyFromMyForm(testDataMessage)
}

function toDateFromStrTime(time: string): Date {
	// '12:05'
	const d = new Date()
	const [HH, mm] = time.split(':')
	d.setHours(parseInt(HH))
	d.setMinutes(parseInt(mm))
	return d
}

const copySheet = () => {
	try {
		// 新シート生成

		// 既存シート数
		const index = ss.getNumSheets()

		// シート名生成
		const fileName = createFileName()

		// シート挿入
		ss.insertSheet(fileName, index + 1)

		// 新しいシートを作成して旧シートからコピーする
		if (sheet !== null) {
			// 最終行
			const lr = sheet?.getLastRow()
			// 最終列
			const lc = sheet?.getLastColumn()
			// 新シート作成
			const newSheet = ss.getSheetByName(fileName)
			// 旧シートからデータを転記
			newSheet?.getRange(1, 1, lr, lc).setValues(sheet?.getRange(1, 1, lr, lc).getValues())

			// おそらくフォーマットが狂うので整形
			// （ここでは7, 8列目に残業開始時間、終了時間が並んでいるものと想定）
			newSheet?.getRange(2, 7, lr - 1, 2).setNumberFormat('hh:mm')
			//（ここでは2,3,6列目の作成日時、変更日時2箇所）
			newSheet?.getRange(2, 2, lr - 1, 2).setNumberFormat('yyyy-mm-dd')
			// （ここでは6列目の実施日）
			newSheet?.getRange(2, 6, lr - 1, 1).setNumberFormat('yyyy-mm-dd')

			// 旧シート初期化
			if (new Date().getDate() === 1) {
				sheet?.deleteRows(2, lr - 1) // あえて.getRange().clear()は使わない
			}
		}

		// トリガーが失敗したら知らせる
	} catch (e) {
		console.log('コピーシートError:: ', e)
	}
}

function setScriptProperty() {
	// ys9d5Voj5gYDXK2GuAbskcpzq77gt2XDs3vMiKwseQB
	scriptProperties.setProperty('LINE_NOTIFY_API_TOKEN', 'ys9d5Voj5gYDXK2GuAbskcpzq77gt2XDs3vMiKwseQB')

	console.log(scriptProperties.getProperty('LINE_NOTIFY_API_TOKEN') ?? '')
}
