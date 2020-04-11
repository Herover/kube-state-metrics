package gdrive

import (
	"encoding/json"
	"fmt"
	"io/ioutil"
	"log"
	"math"
	"net/http"
	"os"
	"strconv"

	"golang.org/x/net/context"
	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
	"google.golang.org/api/drive/v3"
	"google.golang.org/api/sheets/v4"
)

// Retrieve a token, saves the token, then returns the generated client.
func getClient(config *oauth2.Config) *http.Client {
	// The file token.json stores the user's access and refresh tokens, and is
	// created automatically when the authorization flow completes for the first
	// time.
	tokFile := "token.json"
	tok, err := tokenFromFile(tokFile)
	if err != nil {
		tok = getTokenFromWeb(config)
		saveToken(tokFile, tok)
	}
	return config.Client(context.Background(), tok)
}

// Request a token from the web, then returns the retrieved token.
func getTokenFromWeb(config *oauth2.Config) *oauth2.Token {
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

// Retrieves a token from a local file.
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

// Saves a token to a file path.
func saveToken(path string, token *oauth2.Token) {
	fmt.Printf("Saving credential file to: %s\n", path)
	f, err := os.OpenFile(path, os.O_RDWR|os.O_CREATE|os.O_TRUNC, 0600)
	if err != nil {
		log.Fatalf("Unable to cache oauth token: %v", err)
	}
	defer f.Close()
	json.NewEncoder(f).Encode(token)
}

// GDrive will hold services related to google drive services
type GDrive struct {
	driveSrv  *drive.Service
	sheetsSrv *sheets.Service

	// ID of base folder
	folderID string

	// Maps each thing we want a sheet for to a sheet
	sheets map[string]*sheets.Spreadsheet
}

func (gDrive *GDrive) updateSheetHead(fileID, title, description string, data map[string][]interface{}) error {
	keys := make([]interface{}, 0, len(data))
	for k := range data {
		keys = append(keys, k)
	}
	values := [][]interface{}{{description}, keys}
	valueRange := &sheets.ValueRange{
		Values: values,
	}
	_, err := gDrive.sheetsSrv.Spreadsheets.Values.Update(
		fileID,
		"A1:"+columnToLetter(1+len(keys))+strconv.Itoa(len(values)),
		valueRange,
	).
		ValueInputOption("RAW").
		Do()
	if err != nil {
		return err
	}
	return nil
}

func (gDrive *GDrive) getSheet(title string, description string, data map[string][]interface{}) (*sheets.Spreadsheet, error) {
	var sheet *sheets.Spreadsheet
	sheet, exists := gDrive.sheets[title]
	if !exists {
		files, err := gDrive.driveSrv.Files.List().
			Q(
				fmt.Sprintf("name = '%s'", title) +
					" and" +
					fmt.Sprintf("'%s' in parents", gDrive.folderID),
			).
			Do()
		if err != nil {
			return nil, err
		}
		/*
			for _, file := range files.Files {
				err := gDrive.driveSrv.Files.Delete(file.Id).Do()
				if err != nil {
					return nil, err
				}
			}
			return nil, nil
		*/
		if len(files.Files) == 0 {
			file, err := gDrive.driveSrv.Files.Create(&drive.File{
				Name:     title,
				MimeType: "application/vnd.google-apps.spreadsheet",
				Parents:  []string{gDrive.folderID},
			}).Do()
			if err != nil {
				return nil, err
			}
			err = gDrive.updateSheetHead(file.Id, title, description, data)
			if err != nil {
				return nil, err
			}

			sheet, err = gDrive.sheetsSrv.Spreadsheets.Get(file.Id).Do()
			if err != nil {
				return nil, err
			}

		} else if len(files.Files) == 1 {
			err = gDrive.updateSheetHead(files.Files[0].Id, title, description, data)
			if err != nil {
				return nil, err
			}
			sheet, err = gDrive.sheetsSrv.Spreadsheets.Get(files.Files[0].Id).Do()
		} else {
			return nil, fmt.Errorf("Got %d files with name %s, expected 1", len(files.Files), title)
		}
	}

	return sheet, nil
}

func (gDrive GDrive) LogRow(title string, description string, row map[string][]interface{}) error {
	_, err := gDrive.getSheet(title, description, row)
	if err != nil {
		return err
	}
	/*
		_, err = gDrive.sheetsSrv.Spreadsheets.Get(sheet.SpreadsheetId).Do()
		if err != nil {
			return err
		}
	*/
	return nil
}

// Create sets up connections and credentials
func Create() GDrive {
	b, err := ioutil.ReadFile("credentials.json")
	if err != nil {
		log.Fatalf("Unable to read client secret file: %v", err)
	}

	// If modifying these scopes, delete your previously saved token.json.
	config, err := google.ConfigFromJSON(b, "https://www.googleapis.com/auth/drive.file")
	if err != nil {
		log.Fatalf("Unable to parse client secret file to config: %v", err)
	}
	client := getClient(config)

	driveSrv, err := drive.New(client)
	if err != nil {
		log.Fatalf("Unable to retrieve Drive client: %v", err)
	}

	files, err := driveSrv.Files.List().
		Q("mimeType = 'application/vnd.google-apps.folder'").
		Do()
	if err != nil {
		log.Fatalf("Unable to retrieve files: %v", err)
	}

	var folder *drive.File
	for _, file := range files.Files {
		if file.Name == "k8sheets" {
			folder = file
		}
	}
	if folder == nil {
		folder, err = driveSrv.Files.Create(&drive.File{
			Name:     "k8sheets",
			MimeType: "application/vnd.google-apps.folder",
		}).
			Do()
		if err != nil {
			log.Fatalf("Unable to create folder: %v", err)
		}
	}

	sheetsSrv, err := sheets.New(client)
	if err != nil {
		log.Fatalf("Unable to retrieve Sheets client: %v", err)
	}

	gDrive := GDrive{
		driveSrv:  driveSrv,
		sheetsSrv: sheetsSrv,
		folderID:  folder.Id,
	}

	return gDrive
}

// https://stackoverflow.com/a/21231012
func columnToLetter(column int) string {
	temp := 0
	letter := ""
	for column > 0 {
		temp = (column - 1) % 26
		letter = string(temp+65) + letter
		column = (column - temp - 1) / 26
	}
	return letter
}

func letterToColumn(letter string) int {
	column := 0
	length := len(letter)
	for i := 0; i < length; i++ {
		column += (int(letter[i]) - 64) * int(math.Pow(26, float64(length-i-1)))
	}
	return column
}
