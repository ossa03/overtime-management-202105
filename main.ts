//? スプレッドシート 時間外勤務表の管理
const ssID = '1EhXDA7gHDe7hNZqbjDhERZabgYCrSsvIXaL-ohRr2Eg'
const ss: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ssID)
const sheet: GoogleAppsScript.Spreadsheet.Sheet | null = ss.getSheetByName('今月')

function main() {
	console.log('This is Main function !')

	const sheetName = sheet?.getName()
	if (sheetName !== undefined) {
		console.log(sheetName)
	}
}

// TODO,指定日,指定時間,を取得する関数
// function getItemsFromSpreadSheet(): {
// 	created_at: Date
// 	updated_at: Date
// 	start: Date
// 	end: Date
// 	name: string
// } {
// 	// スプレッドシートの最後の行の項目をすべて取得する
// 	const range: GoogleAppsScript.Spreadsheet.Range | undefined = sheet?.getDataRange()
// 	if (range !== undefined) {
// 		const rows = range.getValues()
// 		const lastRow: number = rows.length - 1
// 		console.log('LASTrow:  ', rows[lastRow])

// 		const lastRowValues = rows[lastRow] // [timestamp,todo,date,time,.............]

// 		// const [timestamp, todo, date, time] = lastRowValues
// 		const timestamp: Date = lastRowValues[0] // Dateオブジェクト
// 		const todo = lastRowValues[1]
// 		const date = Utilities.formatDate(lastRowValues[2], 'Asia/Tokyo', 'yyyy-MM-dd')
// 		const time = Utilities.formatDate(lastRowValues[3], 'Asia/Tokyo', 'HH:mm')
// 		// Utilities.formatDate(lastRowValues[2], 'Asia/Tokyo', 'yyyy-MM-dd')
// 		// Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')
// 		console.log('Result: ', todo, date, time)

// 		return { todo, date, time }
// 	}
// 	return false
// }

function lineNotifyFromMyForm() {
	const LINE_NOTIFY_API_TOKEN = PropertiesService.getScriptProperties().getProperty('LINE_NOTIFY_API_TOKEN')
	const LINE_NOTIFY_API_URL = 'https://notify-api.line.me/api/notify'

	// TODOを取得する
	// const { todo } = getItemsFromSpreadSheet()

	const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
		method: 'post',
		headers: {
			'Content-Type': 'application/x-www-form-urlencoded',
			Authorization: `Bearer ${LINE_NOTIFY_API_TOKEN}`,
		},
		payload: `message=\n\n`,
	}

	// LINEに通知
	UrlFetchApp.fetch(LINE_NOTIFY_API_URL, options)
}

//? POSTリクエストが来たときの処理 ---------------------------------------------------
function doPost(e: GoogleAppsScript.Events.DoPost) {
	// デバッグlog
	console.log('e:  ', e)

	// ReactAppからPOSTされたデータを取得する
	const data: OvertimePost = JSON.parse(e.postData.contents)
	let { name, start, end, description } = data

	// タイムスタンプ生成
	const created_at = new Date()

	// デバッグlog
	console.log('created_at:  ', created_at)
	console.log('name:  ', name)
	console.log('start:  ', start)
	console.log('end:  ', end)
	console.log('description:  ', description)

	//TODO スプレッドシートにフォームデータを書き込んで、lineNotify関数を実行する．
	// スプレッドシートに追加する項目↓
	// [ uuid, created_at, updated_at, name, start, end, description]
	const uuid = Utilities.getUuid()

	const addData = sheet?.appendRow([uuid, created_at, created_at, name, start, end, description])

	// スプレッドシートに書き込まれるまで少し待機
	Utilities.sleep(10 * 1000)
	console.log('appended!')

	// LINEへ通知
	lineNotifyFromMyForm()
}
