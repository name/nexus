package main

import (
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"

	"nexus/internal/msi"

	"syscall"

	tea "github.com/charmbracelet/bubbletea"
	"github.com/charmbracelet/lipgloss"
	"github.com/spf13/cobra"
)

var rootCmd = &cobra.Command{
	Use:   "nexus",
	Short: "Nexus - Intune application management tool",
	Run:   run_interactive,
}

var (
	titleStyle = lipgloss.NewStyle().
		Foreground(lipgloss.Color("#FFFFFF")).
		Border(lipgloss.RoundedBorder()).
		BorderForeground(lipgloss.Color("#FF875F")).
		Margin(0, 0, 1, 0).
		Padding(0, 2).
		Width(50)
)

func init() {
	rootCmd.AddCommand(&cobra.Command{
		Use:   "config",
		Short: "Configure Nexus settings",
	})
}

func main() {
	if err := rootCmd.Execute(); err != nil {
		fmt.Println(err)
	}
}

type model struct {
	step          int
	cursor        int
	source        string
	installerType string
	packageName   string
	textInput     string
	typing        bool
	outputDir     string
	err           error
	validationErr string
	productCode   string
}

func validateInput(source, installerType, input string) error {
	if source == "Download from Internet" {
		if !strings.HasPrefix(strings.ToLower(input), "https://") {
			return fmt.Errorf("URL must start with 'https://'")
		}
	} else {
		ext := strings.ToLower(filepath.Ext(input))
		if installerType == "MSI" && ext != ".msi" {
			return fmt.Errorf("file must have .msi extension")
		}
		if installerType == "EXE" && ext != ".exe" {
			return fmt.Errorf("file must have .exe extension")
		}

		if _, err := os.Stat(input); os.IsNotExist(err) {
			return fmt.Errorf("file does not exist: %s", input)
		}
	}
	return nil
}

func (m model) Init() tea.Cmd {
	return nil
}

func (m model) Update(msg tea.Msg) (tea.Model, tea.Cmd) {
	switch msg := msg.(type) {
	case tea.KeyMsg:
		if msg.String() == "ctrl+c" {
			return m, tea.Quit
		}

		if m.step == 2 {
			m.typing = true
			switch msg.Type {
			case tea.KeyEnter:
				if m.packageName == "" {
					m.packageName = m.textInput
					m.textInput = ""
					return m, nil
				}
				if err := validateInput(m.source, m.installerType, m.textInput); err != nil {
					m.validationErr = err.Error()
					return m, nil
				}

				tmp_dir, err := os.MkdirTemp("", "nexus-package-")
				if err != nil {
					m.err = err
					return m, tea.Quit
				}
				m.outputDir = tmp_dir
				m.step++
				m.typing = false
				m.cursor = 0
				m.validationErr = ""
				return m, nil
			case tea.KeyBackspace:
				if len(m.textInput) > 0 {
					m.textInput = m.textInput[:len(m.textInput)-1]
				}
				return m, nil
			case tea.KeyRunes:
				m.textInput += string(msg.Runes)
				return m, nil
			}
		} else if m.step == 3 {
			switch msg.String() {
			case "up", "k":
				if m.cursor > 0 {
					m.cursor--
				}
			case "down", "j":
				if m.cursor < 1 {
					m.cursor++
				}
			case "enter":
				if m.cursor == 0 {
					return m, tea.Quit
				} else {
					os.RemoveAll(m.outputDir)
					m.step = 0
					m.cursor = 0
					m.source = ""
					m.installerType = ""
					m.textInput = ""
					m.typing = false
					m.outputDir = ""
					return m, nil
				}
			}
		} else {
			switch msg.String() {
			case "up", "k":
				if m.cursor > 0 {
					m.cursor--
				}
			case "down", "j":
				switch m.step {
				case 0, 1:
					if m.cursor < 1 {
						m.cursor++
					}
				}
			case "enter":
				switch m.step {
				case 0:
					m.source = []string{"Local File", "Download from Internet"}[m.cursor]
					m.step++
					m.cursor = 0
				case 1:
					m.installerType = []string{"MSI", "EXE"}[m.cursor]
					m.step++
					m.cursor = 0
				}
			}
		}
	}
	return m, nil
}

func getScriptContent(installerType, packageName string) (install, uninstall string) {
	if installerType == "MSI" {
		install = fmt.Sprintf(`# Install script for %s
# This is a placeholder for the MSI installation script
# The actual script will handle:
# - MSI installation with logging
# - Error handling
# - Exit code checking`, packageName)

		uninstall = fmt.Sprintf(`# Uninstall script for %s
# This is a placeholder for the MSI uninstallation script
# The actual script will handle:
# - MSI removal with logging
# - Error handling
# - Exit code checking`, packageName)
	} else {
		install = fmt.Sprintf(`# Install script for %s
# This is a placeholder for the EXE installation script
# The actual script will handle:
# - EXE installation with logging
# - Error handling
# - Exit code checking`, packageName)

		uninstall = fmt.Sprintf(`# Uninstall script for %s
# This is a placeholder for the EXE uninstallation script
# The actual script will handle:
# - EXE removal with logging
# - Error handling
# - Exit code checking`, packageName)
	}
	return
}

func (m model) View() string {
	s := titleStyle.Render("❋ Welcome to Nexus CLI preview!")
	s += "\n\n"

	switch m.step {
	case 0:
		s += "Select Installation Source:\n\n"
		choices := []string{"Local File", "Download from Internet"}
		for i, choice := range choices {
			cursor := " "
			if m.cursor == i {
				cursor = ">"
			}
			s += fmt.Sprintf("%s %s\n", cursor, choice)
		}
	case 1:
		s += "Select Installer Type:\n\n"
		choices := []string{"MSI", "EXE"}
		for i, choice := range choices {
			cursor := " "
			if m.cursor == i {
				cursor = ">"
			}
			s += fmt.Sprintf("%s %s\n", cursor, choice)
		}
	case 2:
		if m.packageName == "" {
			s += "Enter package name: " + m.textInput
		} else {
			label := "path to"
			if m.source == "Download from Internet" {
				label = "download URL for"
			}
			s += fmt.Sprintf("Enter %s %s file: %s", label, m.installerType, m.textInput)
		}

		if m.validationErr != "" {
			s += "\n\n" + lipgloss.NewStyle().
				Foreground(lipgloss.Color("#FF0000")).
				Render("Error: "+m.validationErr)
		}
	case 3:
		s += titleStyle.Render("Package Summary")
		s += "\n"
		indent := "    "

		s += lipgloss.NewStyle().Bold(true).Render("Package Details")
		s += fmt.Sprintf("\n%s• Name: %s\n", indent, m.packageName)
		s += fmt.Sprintf("%s• Type: %s\n", indent, m.installerType)

		s += lipgloss.NewStyle().Bold(true).Render("\nSource")
		s += fmt.Sprintf("\n%s• Type: %s\n", indent, m.source)
		if m.source == "Local File" {
			s += fmt.Sprintf("%s• Path: %s\n", indent, m.textInput)
		} else {
			s += fmt.Sprintf("%s• URL: %s\n", indent, m.textInput)
		}

		s += lipgloss.NewStyle().Bold(true).Render("\nIntune Commands")
		s += "\n"
		install, uninstall := getScriptContent(m.installerType, m.packageName)
		s += fmt.Sprintf("%s• Install Script:\n%s  %s\n", indent, indent, strings.ReplaceAll(install, "\n", "\n"+indent+"  "))
		s += fmt.Sprintf("\n%s• Uninstall Script:\n%s  %s\n", indent, indent, strings.ReplaceAll(uninstall, "\n", "\n"+indent+"  "))

		s += lipgloss.NewStyle().Bold(true).Render("\nConfirmation")
		s += "\n"
		choices := []string{"Yes, create package", "No, start over"}
		for i, choice := range choices {
			cursor := fmt.Sprintf("%s  ", indent)
			if m.cursor == i {
				cursor = fmt.Sprintf("%s▸ ", indent)
			}
			s += fmt.Sprintf("%s%s\n", cursor, choice)
		}
	}

	s += "\n" + lipgloss.NewStyle().Faint(true).Render("Press Ctrl+C to quit")
	return s
}

func copyFileToDir(source, destDir, filename string) error {
	sourceFile, err := os.Open(source)
	if err != nil {
		return fmt.Errorf("failed to open source file: %v", err)
	}
	defer sourceFile.Close()

	destPath := filepath.Join(destDir, filepath.Base(filename))
	destFile, err := os.Create(destPath)
	if err != nil {
		return fmt.Errorf("failed to create destination file: %v", err)
	}
	defer destFile.Close()

	if _, err := io.Copy(destFile, sourceFile); err != nil {
		return fmt.Errorf("failed to copy file: %v", err)
	}

	return nil
}

func getMSIProductCode(msiPath string) (string, error) {
	msiPathW, err := syscall.UTF16PtrFromString(msiPath)
	if err != nil {
		return "", fmt.Errorf("failed to convert path: %v", err)
	}

	var handle syscall.Handle
	persist, _ := syscall.UTF16PtrFromString("0")
	if err := msi.OpenDatabase(msiPathW, persist, &handle); err != nil {
		return "", fmt.Errorf("failed to open MSI database: %v", err)
	}
	defer msi.CloseHandle(handle)

	var view syscall.Handle
	query := "SELECT `Value` FROM `Property` WHERE `Property` = 'ProductCode'"
	queryW, _ := syscall.UTF16PtrFromString(query)
	if err := msi.DatabaseOpenView(handle, queryW, &view); err != nil {
		return "", fmt.Errorf("failed to open view: %v", err)
	}
	defer msi.CloseHandle(view)

	if err := msi.ViewExecute(view, 0); err != nil {
		return "", fmt.Errorf("failed to execute view: %v", err)
	}

	var record syscall.Handle
	if err := msi.ViewFetch(view, &record); err != nil {
		return "", fmt.Errorf("failed to fetch record: %v", err)
	}
	defer msi.CloseHandle(record)

	var buf [39]uint16
	bufLen := uint32(len(buf))
	if err := msi.RecordGetString(record, 1, &buf[0], &bufLen); err != nil {
		return "", fmt.Errorf("failed to get product code: %v", err)
	}

	return syscall.UTF16ToString(buf[:]), nil
}

func run_interactive(cmd *cobra.Command, args []string) {
	p := tea.NewProgram(model{})
	m, err := p.Run()
	if err != nil {
		fmt.Printf("Error running program: %v\n", err)
		return
	}

	finalModel := m.(model)
	if finalModel.step == 3 && finalModel.cursor == 0 {
		fmt.Println("\n" + titleStyle.Render("Creating Package"))
		indent := "    "
		sectionStyle := lipgloss.NewStyle().Bold(true)

		intunewinFile := fmt.Sprintf("%s.intunewin", finalModel.packageName)
		installerFile := fmt.Sprintf("%s.%s",
			strings.TrimSpace(finalModel.packageName),
			strings.ToLower(finalModel.installerType))

		if finalModel.source == "Local File" {
			fmt.Printf("%s• Copying installer file...\n", indent)
			if err := copyFileToDir(finalModel.textInput, finalModel.outputDir, installerFile); err != nil {
				fmt.Printf("Error copying installer: %v\n", err)
				return
			}

			if finalModel.installerType == "MSI" {
				fmt.Printf("%s• Extracting MSI product code...\n", indent)
				msiPath := filepath.Join(finalModel.outputDir, installerFile)
				productCode, err := getMSIProductCode(msiPath)
				if err != nil {
					fmt.Printf("Error getting product code: %v\n", err)
					return
				}
				finalModel.productCode = productCode
				fmt.Printf("%s• MSI Product Code: %s\n", indent, finalModel.productCode)
			}
		} else {
			fmt.Printf("%s• Download functionality not yet implemented\n", indent)
			return
		}

		fmt.Println("\n" + sectionStyle.Render("Summary"))
		fmt.Printf("%s• Creating %s package: %s\n", indent, finalModel.installerType, finalModel.packageName)
		if finalModel.source == "Download from Internet" {
			fmt.Printf("%s• Source: %s\n", indent, finalModel.textInput)
		} else {
			fmt.Printf("%s• Source: %s\n", indent, finalModel.textInput)
		}
		fmt.Printf("%s• Output: %s/%s\n", indent, finalModel.outputDir, intunewinFile)

		fmt.Println("\n" + sectionStyle.Render("Intune Configuration"))
		install, uninstall := getScriptContent(finalModel.installerType, finalModel.packageName)

		installPath := filepath.Join(finalModel.outputDir, "install.ps1")
		if err := os.WriteFile(installPath, []byte(install), 0644); err != nil {
			fmt.Printf("Error writing install script: %v\n", err)
			return
		}

		uninstallPath := filepath.Join(finalModel.outputDir, "uninstall.ps1")
		if err := os.WriteFile(uninstallPath, []byte(uninstall), 0644); err != nil {
			fmt.Printf("Error writing uninstall script: %v\n", err)
			return
		}

		fmt.Printf("%s• Install Script:   %s\n", indent, installPath)
		fmt.Printf("%s• Uninstall Script: %s\n", indent, uninstallPath)

		fmt.Println("\n" + sectionStyle.Render("Intune Commands"))
		fmt.Printf("%s• Install Command:\n", indent)
		fmt.Printf("%s  powershell.exe -ExecutionPolicy Bypass -File .\\install.ps1\n", indent)
		fmt.Printf("%s• Uninstall Command:\n", indent)
		fmt.Printf("%s  powershell.exe -ExecutionPolicy Bypass -File .\\uninstall.ps1\n", indent)

		fmt.Println("\n" + sectionStyle.Render("Detection Rules"))
		if finalModel.installerType == "MSI" {
			fmt.Printf("%s• Use MSI Product Code for detection: %s\n", indent, finalModel.productCode)
		} else {
			fmt.Printf("%s• Use custom detection script or file existence\n", indent)
		}
	}
}
