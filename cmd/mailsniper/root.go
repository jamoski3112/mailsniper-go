package cmd

import (
	"fmt"
	"os"

	"github.com/spf13/cobra"
)

const banner = `
  __  __       _ _ ____        _                  
 |  \/  | __ _(_) / ___| _ __ (_)_ __   ___ _ __  
 | |\/| |/ _' | | \___ \| '_ \| | '_ \ / _ \ '__| 
 | |  | | (_| | | |___) | | | | | |_) |  __/ |    
 |_|  |_|\__,_|_|_|____/|_| |_|_| .__/ \___|_|    
                                  |_|               

`

// NewRootCmd builds the root cobra command with all subcommands registered.
func NewRootCmd() *cobra.Command {
	rootCmd := &cobra.Command{
		Use:          "mailsniper",
		Short:        "MailSniper - Exchange/O365 email search and recon tool (Go port)",
		Long:         banner + "\nMailSniper is a penetration testing tool for searching email in Microsoft Exchange environments.",
		SilenceUsage: true,
	}

	rootCmd.AddCommand(
		newSelfSearchCmd(),
		newGlobalSearchCmd(),
		newGetGALCmd(),
		newSprayOWACmd(),
		newSprayEWSCmd(),
		newHarvestUsersCmd(),
		newHarvestDomainCmd(),
		newOpenInboxCmd(),
		newListFoldersCmd(),
		newGetADUserCmd(),
		newSendEmailCmd(),
	)

	return rootCmd
}

// Execute runs the root command.
func Execute() {
	if err := NewRootCmd().Execute(); err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
}
