package cmd

import (
	"fmt"
	"github.com/spf13/cobra"
	"github.com/sschaenz/XLSlite/model"
	"os"
)

var (
	source     string
	sheet      string
	output     string
	numRows    int
	headerRow  int
	contentRow int
)

func init() {
	importCmd := &cobra.Command{
		Use:   "import",
		Short: "Importiere Excel-Datei in SQLite-Datenbank",
		Long:  `import —source <name der excel-Datei> —-sheet <name des Tabellenblatts> —-header-row <number of row> —content-row <zeile mit Beginn der Daten> —numOfRows <Anzahl der Zeilen mit Daten> —output <name der sqlite-datei>`,
		Run: func(cmd *cobra.Command, args []string) {
			model.ExcelToDB(source, sheet, headerRow, contentRow, numRows, output)
		},
	}

	importCmd.Flags().StringVarP(&source, "source", "s", "", "Name der Excel-Datei (erforderlich)")
	importCmd.Flags().StringVarP(&sheet, "sheet", "t", "", "Name des Tabellenblatts (erforderlich)")
	importCmd.Flags().IntVarP(&headerRow, "header-row", "c", 1, "Zeile des Tabellen-Headers (optional, Standard: 1)")
	importCmd.Flags().IntVarP(&contentRow, "content-row", "b", 2, "Zeile, in der die Daten beginnen (optional, Standard: Zeile nach Header)")
	importCmd.Flags().IntVarP(&numRows, "numOfRows", "n", -1, "Anzahl der Zeilen, die importiert werden sollen (optional, Standard: alle)")
	importCmd.Flags().StringVarP(&output, "output", "o", "", "Name der SQLite-Datei (erforderlich)")

	importCmd.MarkFlagRequired("source")
	importCmd.MarkFlagRequired("sheet")
	importCmd.MarkFlagRequired("output")

	var rootCmd = &cobra.Command{Use: "EXSqlite"}
	rootCmd.AddCommand(importCmd)
	err := rootCmd.Execute()
	if err != nil {
		fmt.Println(err)
		os.Exit(1)
	}
}
