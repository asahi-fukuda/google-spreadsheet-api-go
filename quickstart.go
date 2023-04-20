package main

import (
	"context"
	"encoding/json"
	"fmt"
	"log"
	"net/http"
	"os"
	"strings"
	"time"

	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
)

// トークンを取得して保存し、生成されたクライアントを返す
func getClient(config *oauth2.Config) *http.Client {
	// token.json ファイルは、ユーザーのアクセスとリフレッシュトークンを保存するファイルで、認証フローが初めて完了したときに自動的に作成
	tokFile := "token.json"
	tok, err := tokenFromFile(tokFile)
	if err != nil {
		tok = getTokenFromWeb(config)
		saveToken(tokFile, tok)
	}
	return config.Client(context.Background(), tok)
}

// Webからトークンを要求し、取得したトークンを返す
func getTokenFromWeb(config *oauth2.Config) *oauth2.Token {
	// 認証コードを取得するためのURLを作成
	authURL := config.AuthCodeURL("state-token", oauth2.AccessTypeOffline)
	fmt.Printf("Go to the following link in your browser then type the "+
		"authorization code: \n%v\n", authURL)

	var authCode string
	if _, err := fmt.Scan(&authCode); err != nil {
		log.Fatalf("Unable to read authorization code: %v", err)
	}

	tok, err := config.Exchange(context.TODO(), authCode)
	if err != nil {
		log.Fatalf("Unable to retrieve token from web: %v", err)
	}
	return tok
}

// ローカルファイルからトークンを取得
func tokenFromFile(file string) (*oauth2.Token, error) {
	f, err := os.Open(file)
	if err != nil {
		return nil, err
	}
	defer f.Close()
	tok := &oauth2.Token{}
	err = json.NewDecoder(f).Decode(tok)
	return tok, err
}

// トークンをファイルパスに保存
func saveToken(path string, token *oauth2.Token) {
	fmt.Printf("Saving credential file to: %s\n", path)
	f, err := os.OpenFile(path, os.O_RDWR|os.O_CREATE|os.O_TRUNC, 0600)
	if err != nil {
		log.Fatalf("Unable to cache oauth token: %v", err)
	}
	defer f.Close()
	json.NewEncoder(f).Encode(token)
}

func main() {
	ctx := context.Background()
	b, err := os.ReadFile("credentials.json")
	if err != nil {
		log.Fatalf("Unable to ReaFile: %v", err)
	}

	config, err := google.ConfigFromJSON(b, "https://www.googleapis.com/auth/spreadsheets")
	if err != nil {
		log.Fatalf("Unable to ConfigFromJSON: %v", err)
	}
	client := getClient(config)

	srv, err := sheets.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		log.Fatalf("Unable to NewService: %v", err)
	}

	// 新規スプレッドシートを作成
	spreadsheet := &sheets.Spreadsheet{
		Properties: &sheets.SpreadsheetProperties{
			Title: "勤務表作成テスト",
		},
	}

	newSheet, err := srv.Spreadsheets.Create(spreadsheet).Do()
	if err != nil {
		log.Fatalf("Unable to Create: %v", err)
	}

	// コピー元のID
	sourceSpreadsheetId := ""

	// コピー先のID
	destinationSpreadsheetId := newSheet.SpreadsheetId

	// コピー元のスプレッドシートを取得
	sourceSpreadsheet, err := srv.Spreadsheets.Get(sourceSpreadsheetId).Do()
	if err != nil {
		log.Fatalf("Unable to Get source spreadsheet: %v", err)
	}

	// コピー元のすべてのシートをコピー先にコピー
	for _, sheet := range sourceSpreadsheet.Sheets {
		rb := &sheets.CopySheetToAnotherSpreadsheetRequest{
			DestinationSpreadsheetId: destinationSpreadsheetId,
		}

		resp, err := srv.Spreadsheets.Sheets.CopyTo(sourceSpreadsheetId, sheet.Properties.SheetId, rb).Context(ctx).Do()
		if err != nil {
			log.Fatal(err)
		}

		// コピー後のシート名から「のコピー」を削除
		newSheetTitle := strings.TrimSuffix(resp.Title, "のコピー")

		updateSheetNameRequest := sheets.Request{
			UpdateSheetProperties: &sheets.UpdateSheetPropertiesRequest{
				Properties: &sheets.SheetProperties{
					SheetId: resp.SheetId,
					Title:   newSheetTitle,
				},
				Fields: "title",
			},
		}

		batchUpdateRequest := &sheets.BatchUpdateSpreadsheetRequest{
			Requests: []*sheets.Request{&updateSheetNameRequest},
		}

		// シート名を更新
		_, err = srv.Spreadsheets.BatchUpdate(destinationSpreadsheetId, batchUpdateRequest).Context(ctx).Do()
		if err != nil {
			log.Fatalf("Unable to update sheet name: %v", err)
		}
	}

	// 空白のスプレッドシートのID
	blankSheetId := newSheet.Sheets[0].Properties.SheetId

	// 空白シートの削除
	if len(sourceSpreadsheet.Sheets) > 0 {
		deleteSheetRequest := sheets.Request{
			DeleteSheet: &sheets.DeleteSheetRequest{
				SheetId: blankSheetId,
			},
		}

		batchUpdateRequest := &sheets.BatchUpdateSpreadsheetRequest{
			Requests: []*sheets.Request{&deleteSheetRequest},
		}

		_, err = srv.Spreadsheets.BatchUpdate(destinationSpreadsheetId, batchUpdateRequest).Context(ctx).Do()
		if err != nil {
			log.Fatalf("Unable to delete blank sheet: %v", err)
		}
	}

	// コピーしたシートを取得
	s, err := srv.Spreadsheets.Get(destinationSpreadsheetId).Do()
	if err != nil {
		log.Fatalf("Unable to retrieve sheets: %v", err)
	}

	now := time.Now()
	year := now.Year()
	month := int(now.Month())

	// 今年と今月の数値をA1とA3に書き込む
	for _, sheet := range s.Sheets {
		sheetName := sheet.Properties.Title
		values := [][]interface{}{
			{year},
			{},
			{month},
		}

		updateValuesRequest := &sheets.ValueRange{
			Range:          sheetName + "!A1:A3",
			Values:         values,
			MajorDimension: "ROWS",
		}

		_, err := srv.Spreadsheets.Values.Update(destinationSpreadsheetId, updateValuesRequest.Range, updateValuesRequest).ValueInputOption("RAW").Context(ctx).Do()
		if err != nil {
			log.Fatalf("Unable to update cells with year and month: %v", err)
		}
	}
}
