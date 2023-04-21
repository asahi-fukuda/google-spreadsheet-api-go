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

// スプレッドシートの新規作成
func createSpreadsheet(srv *sheets.Service) (*sheets.Spreadsheet, error) {
	spreadsheet := &sheets.Spreadsheet{
		Properties: &sheets.SpreadsheetProperties{
			Title: "勤務表作成テスト",
		},
	}

	newSheet, err := srv.Spreadsheets.Create(spreadsheet).Do()
	if err != nil {
		return nil, err
	}

	return newSheet, nil
}

// スプレッドシートをシートIDから取得
func getSpreadsheet(srv *sheets.Service, spreadsheetId string) (*sheets.Spreadsheet, error) {
	spreadsheet, err := srv.Spreadsheets.Get(spreadsheetId).Do()
	if err != nil {
		return nil, err
	}

	return spreadsheet, nil
}

// IDで指定したスプレッドシートをコピー
func copySpreadsheet(ctx context.Context, sourceSpreadsheet *sheets.Spreadsheet, srv *sheets.Service, sourceSpreadsheetId string, destinationSpreadsheetId string) error {
	for _, sheet := range sourceSpreadsheet.Sheets {
		rb := &sheets.CopySheetToAnotherSpreadsheetRequest{
			DestinationSpreadsheetId: destinationSpreadsheetId,
		}

		resp, err := srv.Spreadsheets.Sheets.CopyTo(sourceSpreadsheetId, sheet.Properties.SheetId, rb).Context(ctx).Do()
		if err != nil {
			log.Fatal(err)
		}

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

		_, err = srv.Spreadsheets.BatchUpdate(destinationSpreadsheetId, batchUpdateRequest).Context(ctx).Do()
		if err != nil {
			log.Fatalf("Unable to update sheet name: %v", err)
		}
	}

	return nil
}

// 空白のスプレッドシートを削除
func deleteBlankSheet(ctx context.Context, srv *sheets.Service, newSheet *sheets.Spreadsheet, destinationSpreadsheetId string) error {
	blankSheetId := newSheet.Sheets[0].Properties.SheetId

	deleteSheetRequest := sheets.Request{
		DeleteSheet: &sheets.DeleteSheetRequest{
			SheetId: blankSheetId,
		},
	}

	batchUpdateRequest := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{&deleteSheetRequest},
	}

	_, err := srv.Spreadsheets.BatchUpdate(destinationSpreadsheetId, batchUpdateRequest).Context(ctx).Do()
	if err != nil {
		log.Fatalf("Unable to delete sheet: %v", err)
	}

	return nil
}

// セルA1とA3に年と月を入力
func updateCellsYearMonth(ctx context.Context, srv *sheets.Service, destinationSpreadsheet *sheets.Spreadsheet, destinationSpreadsheetId string) error {
	now := time.Now()
	year := now.Year()
	month := int(now.Month())

	for _, sheet := range destinationSpreadsheet.Sheets {
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

	return nil
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

	newSheet, err := createSpreadsheet(srv)
	if err != nil {
		log.Fatalf("Unable to createSpreadsheet: %v", err)
	}

	// コピー元のID
	sourceSpreadsheetId := ""

	// コピー先のID（作成したID）
	destinationSpreadsheetId := newSheet.SpreadsheetId

	sourceSpreadsheet, err := getSpreadsheet(srv, sourceSpreadsheetId)
	if err != nil {
		log.Fatalf("Unable to Get source spreadsheet: %v", err)
	}

	err = copySpreadsheet(ctx, sourceSpreadsheet, srv, sourceSpreadsheetId, destinationSpreadsheetId)
	if err != nil {
		log.Fatalf("Unable to copySpreadsheet: %v", err)
	}

	if len(sourceSpreadsheet.Sheets) > 0 {
		err = deleteBlankSheet(ctx, srv, newSheet, destinationSpreadsheetId)
		if err != nil {
			log.Fatalf("Unable to delete blank sheet: %v", err)
		}
	}

	destinationSpreadsheet, err := getSpreadsheet(srv, destinationSpreadsheetId)
	if err != nil {
		log.Fatalf("Unable to retrieve sheets: %v", err)
	}

	err = updateCellsYearMonth(ctx, srv, destinationSpreadsheet, destinationSpreadsheetId)
	if err != nil {
		log.Fatalf("Unable to update cells with year and month: %v", err)
	}
}
