package cmd

import (
	"github.com/spf13/cobra"
)

var rootCmd = &cobra.Command{
	Use:   "XLSqlite",
	Short: "Eine CLI-App zum Extrahieren von Daten aus einer Excel-Datei und Importieren in eine SQLite-Datenbank.",
}

func Execute() {
	if err := rootCmd.Execute(); err != nil {
		panic(err)
	}
}
