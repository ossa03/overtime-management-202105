// スプレッドシート 時間外勤務表の管理
const ssID = '1EhXDA7gHDe7hNZqbjDhERZabgYCrSsvIXaL-ohRr2Eg'
const ss: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ssID)
const sheet: GoogleAppsScript.Spreadsheet.Sheet | null = ss.getSheetByName('今月')
const scriptProperties: GoogleAppsScript.Properties.Properties = PropertiesService.getScriptProperties()

function main() {
	console.log('This is Main function !')
	const fileName = createFileName()
	const pdfBlob = createPdfBlob(ss, fileName)
	const pdfFile = createPdfFile(pdfBlob)
	const fileUrl = getFileUrl(pdfFile)

	// メールに送信
	sendEmail(pdfBlob, fileUrl)

	// LINEに送信
	const lineNotifyMessage = `\n\n今月の時間外勤務表\n\n${fileUrl}`
	sendLineNotify(lineNotifyMessage)
}

function sendLineNotify(message: string) {
	const LINE_NOTIFY_API_TOKEN = scriptProperties.getProperty('LINE_NOTIFY_API_TOKEN')
	const LINE_NOTIFY_API_URL = 'https://notify-api.line.me/api/notify'

	// TODOを取得する
	// const { todo } = getItemsFromSpreadSheet()

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
	const parsed_start = parseTime(start, parsed_date)
	const parsed_end = parseTime(end, parsed_date)
	const diff_time = (parsed_end.getTime() - parsed_start.getTime()) / 1000 / 60 // 分

	// 10列目に残業時間のヘッダーが無ければ 入力する．
	const total_time_cell = sheet?.getRange(1, 10)
	if (total_time_cell?.getValue() === '') {
		total_time_cell.setValue('残業時間（分）')
	}

	// spreadsheetに追加する．
	sheet?.appendRow([uuid, created_at, created_at, radiologist, modality, date, start, end, description, diff_time])

	// スプレッドシートに書き込まれるまで少し待機
	Utilities.sleep(5 * 1000)

	// const testDataMessage = `これはテストデータです．\n\n${e.postData.contents}`
	const postMessage = `時間外勤務を登録しました．\n\n${date}\n${radiologist}\n${modality}\n${start} 〜 ${end}\n${description}`
	// LINEへ通知
	sendLineNotify(postMessage)
	// lineNotifyFromMyForm(testDataMessage)
}

function parseTime(stringTime: string, date = new Date()): Date {
	// '12:05'
	const [HH, mm] = stringTime.split(':')
	date.setHours(parseInt(HH))
	date.setMinutes(parseInt(mm))
	return date
}

function setScriptProperty() {
	// ys9d5Voj5gYDXK2GuAbskcpzq77gt2XDs3vMiKwseQB
	scriptProperties.setProperty('LINE_NOTIFY_API_TOKEN', 'ys9d5Voj5gYDXK2GuAbskcpzq77gt2XDs3vMiKwseQB')

	console.log(scriptProperties.getProperty('LINE_NOTIFY_API_TOKEN') ?? '')
}
