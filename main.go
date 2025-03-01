package main

import (
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"os"
	"os/exec"
	"path"
	"path/filepath"
	"regexp"
	"runtime"
	"strings"

	"nexus/internal/msi"

	"syscall"

	_ "embed"

	"github.com/charmbracelet/bubbles/help"
	"github.com/charmbracelet/bubbles/key"
	"github.com/charmbracelet/bubbles/textinput"
	tea "github.com/charmbracelet/bubbletea"
	"github.com/charmbracelet/lipgloss"
	"github.com/spf13/cobra"
	"golang.org/x/text/cases"
	"golang.org/x/text/language"
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

const (
	nexusDir       = "C:\\ProgramData\\Nexus"
	intuneUtilDir  = "C:\\ProgramData\\Nexus\\Tools"
	intuneUtilPath = "C:\\ProgramData\\Nexus\\Tools\\IntuneWinAppUtil.exe"
	intuneUtilUrl  = "https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/raw/master/IntuneWinAppUtil.exe"
	packagesDir    = "C:\\ProgramData\\Nexus\\Packages"
	downloadsDir   = "C:\\ProgramData\\Nexus\\Downloads"
)

//go:embed "templates/Install-Script.ps1"
var installScriptTemplate string

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
	version       string
	mode          string
	packages      []string
	packages_dir  string
	text_input    textinput.Model
	help          help.Model
	keymap        keymap
	suggestions   []string
}

type keymap struct{}

func (k keymap) ShortHelp() []key.Binding {
	return []key.Binding{
		key.NewBinding(key.WithKeys("tab"), key.WithHelp("tab", "complete")),
		key.NewBinding(key.WithKeys("ctrl+c"), key.WithHelp("ctrl+c", "quit")),
	}
}

func (k keymap) FullHelp() [][]key.Binding {
	return [][]key.Binding{k.ShortHelp()}
}

func initial_model() model {
	m := model{
		step:         -1,
		packages_dir: packagesDir,
	}

	// Initialize text input
	ti := textinput.New()
	ti.Placeholder = "application name"
	ti.Prompt = "Package name: "
	ti.PromptStyle = lipgloss.NewStyle().Foreground(lipgloss.Color("#FF875F"))
	ti.Cursor.Style = lipgloss.NewStyle().Foreground(lipgloss.Color("#FF875F"))
	ti.Focus()
	ti.CharLimit = 50
	ti.Width = 30
	ti.ShowSuggestions = true

	// Initialize help
	h := help.New()

	m.text_input = ti
	m.help = h
	m.keymap = keymap{}

	return m
}

func validateInput(source, installerType, input string) error {
	if source == "Download from Internet" {
		if !strings.HasPrefix(strings.ToLower(input), "https://") {
			return fmt.Errorf("URL must start with 'https://'")
		}
	} else {
		cleanPath := filepath.Clean(input)
		absPath, err := filepath.Abs(cleanPath)
		if err != nil {
			return fmt.Errorf("invalid file path: %v", err)
		}

		ext := strings.ToLower(filepath.Ext(absPath))
		if installerType == "MSI" && ext != ".msi" {
			return fmt.Errorf("file must have .msi extension")
		}
		if installerType == "EXE" && ext != ".exe" {
			return fmt.Errorf("file must have .exe extension")
		}

		if _, err := os.Stat(absPath); os.IsNotExist(err) {
			return fmt.Errorf("file does not exist: %s", input)
		}
	}
	return nil
}

func (m model) Init() tea.Cmd {
	// Load existing packages for suggestions
	packages, _ := get_existing_packages(m.packages_dir)
	var suggestions []string
	for _, pkg := range packages {
		suggestions = append(suggestions, pkg)
	}

	// Add some common application names as suggestions
	common_apps := []string{
		"Microsoft Office",
		"Adobe Acrobat Reader",
		"Google Chrome",
		"Mozilla Firefox",
		"Zoom",
		"Microsoft Teams",
		"VLC Media Player",
		"7-Zip",
		"Notepad++",
		"Visual Studio Code",
	}

	for _, app := range common_apps {
		if !contains(suggestions, app) {
			suggestions = append(suggestions, app)
		}
	}

	m.suggestions = suggestions
	m.text_input.SetSuggestions(suggestions)

	return textinput.Blink
}

func contains(slice []string, item string) bool {
	for _, s := range slice {
		if s == item {
			return true
		}
	}
	return false
}

// Add this function to handle file path suggestions
func get_path_suggestions(current_path string) []string {
	var suggestions []string

	// Get the directory to look in
	dir_path := filepath.Dir(current_path)
	if dir_path == "." {
		dir_path = "."
		current_path = ""
	}

	// If empty, start with drives on Windows or root on other OS
	if dir_path == "." && current_path == "" {
		if runtime.GOOS == "windows" {
			for _, drive := range "ABCDEFGHIJKLMNOPQRSTUVWXYZ" {
				drive_path := string(drive) + ":\\"
				if _, err := os.Stat(drive_path); err == nil {
					suggestions = append(suggestions, drive_path)
				}
			}
			return suggestions
		} else {
			dir_path = "/"
		}
	}

	// Read directory contents
	files, err := os.ReadDir(dir_path)
	if err != nil {
		return suggestions
	}

	// Get the base name to filter with
	base_name := filepath.Base(current_path)
	if base_name == "." || dir_path == current_path {
		base_name = ""
	}

	// Add matching entries
	for _, file := range files {
		name := file.Name()
		if strings.HasPrefix(strings.ToLower(name), strings.ToLower(base_name)) {
			full_path := filepath.Join(dir_path, name)
			if file.IsDir() {
				full_path = filepath.Join(full_path, "") // Add trailing separator
			}
			suggestions = append(suggestions, full_path)
		}
	}

	return suggestions
}

func (m model) Update(msg tea.Msg) (tea.Model, tea.Cmd) {
	switch msg := msg.(type) {
	case tea.KeyMsg:
		if msg.String() == "ctrl+c" {
			return m, tea.Quit
		}

		// Handle custom directory input
		if m.step == -2 && m.typing {
			switch msg.Type {
			case tea.KeyEnter:
				input := m.text_input.Value()
				if input == "" {
					input = packagesDir
				}

				if err := os.MkdirAll(input, 0755); err != nil {
					m.validationErr = fmt.Sprintf("Failed to create directory: %v", err)
					return m, nil
				}

				if err := save_config(input); err != nil {
					m.validationErr = fmt.Sprintf("Failed to save configuration: %v", err)
					return m, nil
				}

				m.packages_dir = input
				m.step = -1
				m.cursor = 0
				m.typing = false
				m.text_input.Reset()
				m.validationErr = ""
				return m, nil
			default:
				// If Tab is pressed, update suggestions with directory paths
				if msg.String() == "tab" {
					suggestions := get_path_suggestions(m.text_input.Value())
					m.text_input.SetSuggestions(suggestions)
				}

				var cmd tea.Cmd
				m.text_input, cmd = m.text_input.Update(msg)
				return m, cmd
			}
		}

		// Handle autocomplete input when in package name input step
		if m.step == 2 && m.mode != "Repackage Application" && m.packageName == "" {
			switch msg.Type {
			case tea.KeyEnter:
				m.packageName = m.text_input.Value()
				m.text_input.Reset()
				m.text_input.Focus() // Keep focus for the next input
				return m, nil
			default:
				var cmd tea.Cmd
				m.text_input, cmd = m.text_input.Update(msg)
				return m, cmd
			}
		}

		// Handle file path input with autocomplete
		if m.step == 2 && m.mode != "Repackage Application" && m.packageName != "" {
			switch msg.Type {
			case tea.KeyEnter:
				input := m.text_input.Value()
				if err := validateInput(m.source, m.installerType, input); err != nil {
					m.validationErr = err.Error()
					return m, nil
				}

				m.textInput = input

				sanitized_name := sanitize_package_name(m.packageName)
				package_dir := filepath.Join(m.packages_dir, sanitized_name)

				// Check for existing package directory
				if _, err := os.Stat(package_dir); err == nil {
					entries, err := os.ReadDir(filepath.Dir(package_dir))
					if err == nil {
						prefix := filepath.Base(package_dir)
						for _, entry := range entries {
							if strings.HasPrefix(entry.Name(), prefix) {
								os.RemoveAll(filepath.Join(filepath.Dir(package_dir), entry.Name()))
							}
						}
					}
					if err := os.RemoveAll(package_dir); err != nil {
						m.err = fmt.Errorf("failed to remove existing package directory: %v", err)
						return m, tea.Quit
					}
				}

				if err := os.MkdirAll(package_dir, 0755); err != nil {
					m.err = err
					return m, tea.Quit
				}

				m.outputDir = package_dir
				m.step++
				m.typing = false
				m.cursor = 0
				m.validationErr = ""
				m.text_input.Blur()
				return m, nil
			default:
				// Update suggestions when typing
				if msg.String() == "tab" {
					suggestions := get_path_suggestions(m.text_input.Value())
					m.text_input.SetSuggestions(suggestions)
				}

				var cmd tea.Cmd
				m.text_input, cmd = m.text_input.Update(msg)
				return m, cmd
			}
		}

		if m.step == -2 && m.typing {
			switch msg.Type {
			case tea.KeyEnter:
				if m.textInput == "" {
					m.textInput = packagesDir
				}

				if err := os.MkdirAll(m.textInput, 0755); err != nil {
					m.validationErr = fmt.Sprintf("Failed to create directory: %v", err)
					return m, nil
				}

				if err := save_config(m.textInput); err != nil {
					m.validationErr = fmt.Sprintf("Failed to save configuration: %v", err)
					return m, nil
				}

				m.packages_dir = m.textInput
				m.step = -1
				m.cursor = 0
				m.typing = false
				m.textInput = ""
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
		} else if m.step == -2 {
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
					m.packages_dir = packagesDir
					m.step = -1
					m.cursor = 0
				} else {
					m.typing = true
					m.textInput = packagesDir
				}
			}
		} else if m.step == 2 && m.mode != "Repackage Application" {
			m.typing = true
			switch msg.Type {
			case tea.KeyEnter:
				if m.packageName == "" {
					m.packageName = m.textInput
					m.textInput = ""

					if m.mode == "Repackage Application" {
						sanitized_name := sanitize_package_name(m.packageName)
						package_dir := filepath.Join(m.packages_dir, sanitized_name)

						if _, err := os.Stat(package_dir); os.IsNotExist(err) {
							m.validationErr = fmt.Sprintf("Package '%s' does not exist", m.packageName)
							return m, nil
						}

						m.outputDir = package_dir
						m.step = 3
						m.typing = false
						m.cursor = 0
						m.validationErr = ""
						return m, nil
					}

					return m, nil
				}

				if err := validateInput(m.source, m.installerType, m.textInput); err != nil {
					m.validationErr = err.Error()
					return m, nil
				}

				sanitized_name := sanitize_package_name(m.packageName)
				package_dir := filepath.Join(m.packages_dir, sanitized_name)

				if _, err := os.Stat(package_dir); err == nil {
					entries, err := os.ReadDir(filepath.Dir(package_dir))
					if err == nil {
						prefix := filepath.Base(package_dir)
						for _, entry := range entries {
							if strings.HasPrefix(entry.Name(), prefix) {
								os.RemoveAll(filepath.Join(filepath.Dir(package_dir), entry.Name()))
							}
						}
					}
					if err := os.RemoveAll(package_dir); err != nil {
						m.err = fmt.Errorf("failed to remove existing package directory: %v", err)
						return m, tea.Quit
					}
				}

				if err := os.MkdirAll(package_dir, 0755); err != nil {
					m.err = err
					return m, tea.Quit
				}
				m.outputDir = package_dir
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
		} else if m.step == 2 && m.mode == "Repackage Application" {
			switch msg.String() {
			case "up", "k":
				if m.cursor > 0 {
					m.cursor--
				}
			case "down", "j":
				if m.cursor < len(m.packages)-1 {
					m.cursor++
				}
			case "enter":
				if len(m.packages) > 0 {
					m.packageName = m.packages[m.cursor]
					sanitized_name := sanitize_package_name(m.packageName)
					package_dir := filepath.Join(m.packages_dir, sanitized_name)
					m.outputDir = package_dir
					m.step = 3
					m.cursor = 0
					return m, nil
				}
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
					if m.outputDir != "" {
						os.RemoveAll(m.outputDir)
					}
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
				case -1:
					if m.cursor < 2 {
						m.cursor++
					}
				case 0, 1:
					if m.cursor < 1 {
						m.cursor++
					}
				}
			case "enter":
				switch m.step {
				case -1:
					choices := []string{"New Application Package", "Repackage Application", "Set Packages Directory"}
					m.mode = choices[m.cursor]
					if m.mode == "Repackage Application" {
						packages, err := get_existing_packages(m.packages_dir)
						if err != nil {
							m.err = err
							return m, tea.Quit
						}

						if len(packages) == 0 {
							m.validationErr = "No existing packages found"
							return m, nil
						}

						m.packages = packages
						m.step = 2
						m.cursor = 0
					} else if m.mode == "Set Packages Directory" {
						m.step = -2
						m.cursor = 0
					} else {
						m.step = 0
					}
					m.cursor = 0
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

func getScriptContent(installerType, packageName, version string) (install, uninstall string) {
	install = installScriptTemplate
	install = strings.ReplaceAll(install, "<APP_TITLE>", packageName)
	if version == "" {
		version = "1.0"
	}
	install = strings.ReplaceAll(install, "<VERSION>", version)
	install = strings.ReplaceAll(install, "<INSTALLER_TYPE>", installerType)

	defaultArgs := "/qn /norestart"
	if installerType == "EXE" {
		defaultArgs = "/silent"
	}
	install = strings.ReplaceAll(install, "<INSTALL_ARGS>", defaultArgs)

	if installerType == "MSI" {
		uninstall = fmt.Sprintf(`
$company = "Nexus"
$app_title = "%s"
$logging_path = "C:\ProgramData\$company\$app_title"
$script_name = (Get-Item $PSCommandPath).Basename
$log_file = "$logging_path\$script_name.log"

function write_log {
    param ([string]$log_string)
    $timestamp = Get-Date
    $formatted_log = "$timestamp $log_string"
    try {   
        Add-content $log_file -value $formatted_log -ErrorAction SilentlyContinue
    }
    catch {
        Write-Host $formatted_log
    }
}

write_log "Starting uninstall of $app_title"
$product_code = "%s" # To be replaced with actual product code
try {
    $process = Start-Process "msiexec.exe" -ArgumentList "/x $product_code /qn /norestart" -Wait -PassThru
    if ($process.ExitCode -eq 0) {
        write_log "Successfully uninstalled $app_title"
    } else {
        write_log "Uninstall failed with exit code: $($process.ExitCode)"
        exit 1
    }
} catch {
    write_log "Error during uninstall: $($_.Exception.Message)"
    exit 1
}`, packageName, "{PRODUCT_CODE}")
	} else {
		uninstall = fmt.Sprintf(`
$company = "Nexus"
$app_title = "%s"
$logging_path = "C:\ProgramData\$company\$app_title"
$script_name = (Get-Item $PSCommandPath).Basename
$log_file = "$logging_path\$script_name.log"

function write_log {
    param ([string]$log_string)
    $timestamp = Get-Date
    $formatted_log = "$timestamp $log_string"
    try {   
        Add-content $log_file -value $formatted_log -ErrorAction SilentlyContinue
    }
    catch {
        Write-Host $formatted_log
    }
}

write_log "Starting uninstall of $app_title"
$uninstall_path = "%s" # To be replaced with actual uninstall path
try {
    $process = Start-Process $uninstall_path -ArgumentList "/silent" -Wait -PassThru
    if ($process.ExitCode -eq 0) {
        write_log "Successfully uninstalled $app_title"
    } else {
        write_log "Uninstall failed with exit code: $($process.ExitCode)"
        exit 1
    }
} catch {
    write_log "Error during uninstall: $($_.Exception.Message)"
    exit 1
}`, packageName, "C:\\Program Files\\AppName\\uninstall.exe")
	}

	return
}

func createPackageScripts(outputDir, packageName, installerType string) error {
	install, uninstall := getScriptContent(installerType, packageName, "1.0")

	installPath := filepath.Join(outputDir, "Install.ps1")
	if err := os.WriteFile(installPath, []byte(install), 0644); err != nil {
		return fmt.Errorf("failed to create install script: %v", err)
	}

	uninstallPath := filepath.Join(outputDir, "Uninstall.ps1")
	if err := os.WriteFile(uninstallPath, []byte(uninstall), 0644); err != nil {
		return fmt.Errorf("failed to create uninstall script: %v", err)
	}

	return nil
}

func (m model) View() string {
	s := titleStyle.Render("Welcome to the Nexus CLI preview!")
	s += "\n\n"

	switch m.step {
	case -2:
		s += "Set Packages Directory:\n\n"
		if m.typing {
			m.text_input.Prompt = "Enter custom packages directory path: "
			m.text_input.Placeholder = packagesDir
			s += m.text_input.View()

			if m.validationErr != "" {
				s += "\n\n" + lipgloss.NewStyle().
					Foreground(lipgloss.Color("#FF0000")).
					Render("Error: "+m.validationErr)
			}
		} else {
			choices := []string{"Use Default Directory", "Set Custom Directory"}
			for i, choice := range choices {
				cursor := " "
				if m.cursor == i {
					cursor = ">"
				}
				s += fmt.Sprintf("%s %s\n", cursor, choice)
			}
			s += "\n"
			s += fmt.Sprintf("Current packages directory: %s", m.packages_dir)
		}
	case -1:
		s += "Select Operation:\n\n"
		choices := []string{"New Application Package", "Repackage Application", "Set Packages Directory"}
		for i, choice := range choices {
			cursor := " "
			if m.cursor == i {
				cursor = ">"
			}
			s += fmt.Sprintf("%s %s\n", cursor, choice)
		}
		s += "\n"
		s += fmt.Sprintf("Current packages directory: %s", m.packages_dir)
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
		if m.mode == "Repackage Application" {
			s += "Select package to repackage:\n\n"
			if len(m.packages) == 0 {
				s += "No packages found."
			} else {
				for i, pkg := range m.packages {
					cursor := " "
					if m.cursor == i {
						cursor = ">"
					}
					s += fmt.Sprintf("%s %s\n", cursor, pkg)
				}
			}
		} else if m.packageName == "" {
			s += m.text_input.View()
		} else {
			label := "path to"
			if m.source == "Download from Internet" {
				label = "download URL for"
				m.text_input.Placeholder = "https://example.com/installer.exe"
			} else {
				m.text_input.Placeholder = "C:\\path\\to\\installer.exe"
			}

			m.text_input.Prompt = fmt.Sprintf("Enter %s %s file: ", label, m.installerType)
			s += m.text_input.View()

			if m.validationErr != "" {
				s += "\n\n" + lipgloss.NewStyle().
					Foreground(lipgloss.Color("#FF0000")).
					Render("Error: "+m.validationErr)
			}
		}
	case 3:
		s += titleStyle.Render("Package Summary")
		s += "\n"
		indent := "    "

		s += lipgloss.NewStyle().Bold(true).Render("Package Details")
		s += fmt.Sprintf("\n%s• Name: %s\n", indent, m.packageName)

		if m.mode == "Repackage Application" {
			s += fmt.Sprintf("%s• Operation: Repackage Existing Application\n", indent)
			s += fmt.Sprintf("%s• Package Directory: %s\n", indent, m.outputDir)
		} else {
			s += fmt.Sprintf("%s• Type: %s\n", indent, m.installerType)

			s += lipgloss.NewStyle().Bold(true).Render("\nSource")
			s += fmt.Sprintf("\n%s• Type: %s\n", indent, m.source)
			if m.source == "Local File" {
				s += fmt.Sprintf("%s• Path: %s\n", indent, m.textInput)
			} else {
				s += fmt.Sprintf("%s• URL: %s\n", indent, m.textInput)
			}
		}

		s += lipgloss.NewStyle().Bold(true).Render("\nConfirmation")
		s += "\n"

		choices := []string{}
		if m.mode == "Repackage Application" {
			choices = []string{"Yes, repackage application", "No, start over"}
		} else {
			choices = []string{"Yes, create package", "No, start over"}
		}

		for i, choice := range choices {
			cursor := fmt.Sprintf("%s  ", indent)
			if m.cursor == i {
				cursor = fmt.Sprintf("%s▸ ", indent)
			}
			s += fmt.Sprintf("%s%s\n", cursor, choice)
		}
	}

	s += "\n\n" + m.help.View(m.keymap)
	return s
}

func copyFileToDir(source, destDir, filename string) error {
	cleanSource := filepath.Clean(source)
	absSource, err := filepath.Abs(cleanSource)
	if err != nil {
		return fmt.Errorf("invalid source path: %v", err)
	}

	cleanDest := filepath.Clean(destDir)
	absDest, err := filepath.Abs(cleanDest)
	if err != nil {
		return fmt.Errorf("invalid destination path: %v", err)
	}

	sourceFile, err := os.Open(absSource)
	if err != nil {
		return fmt.Errorf("failed to open source file '%s': %v", absSource, err)
	}
	defer sourceFile.Close()

	destPath := filepath.Join(absDest, filename)
	destFile, err := os.Create(destPath)
	if err != nil {
		return fmt.Errorf("failed to create destination file '%s': %v", destPath, err)
	}
	defer destFile.Close()

	if _, err := io.Copy(destFile, sourceFile); err != nil {
		return fmt.Errorf("failed to copy file: %v", err)
	}

	return nil
}

func getMSIProductCode(msiPath string) (string, string, error) {
	msiPathW, err := syscall.UTF16PtrFromString(msiPath)
	if err != nil {
		return "", "", fmt.Errorf("failed to convert path: %v", err)
	}

	var handle syscall.Handle
	persist, _ := syscall.UTF16PtrFromString("0")
	if err := msi.OpenDatabase(msiPathW, persist, &handle); err != nil {
		return "", "", fmt.Errorf("failed to open MSI database: %v", err)
	}
	defer msi.CloseHandle(handle)

	productCode, err := getMSIProperty(handle, "ProductCode")
	if err != nil {
		return "", "", fmt.Errorf("failed to get product code: %v", err)
	}

	version, err := getMSIProperty(handle, "ProductVersion")
	if err != nil {
		return productCode, "", fmt.Errorf("failed to get version: %v", err)
	}

	return productCode, version, nil
}

func getMSIProperty(handle syscall.Handle, property string) (string, error) {
	var view syscall.Handle
	query := fmt.Sprintf("SELECT `Value` FROM `Property` WHERE `Property` = '%s'", property)
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
		return "", fmt.Errorf("failed to get property value: %v", err)
	}

	return syscall.UTF16ToString(buf[:]), nil
}

func run_interactive(cmd *cobra.Command, args []string) {
	if err := ensureNexusDirs(); err != nil {
		fmt.Printf("Error setting up Nexus directories: %v\n", err)
		return
	}

	if err := ensureIntuneUtil(); err != nil {
		fmt.Printf("Error setting up Intune tools: %v\n", err)
		return
	}

	packages_dir, err := load_config()
	if err != nil {
		fmt.Printf("Error loading configuration: %v\n", err)
		packages_dir = packagesDir
	}

	// Use the new initial_model function
	m := initial_model()
	m.packages_dir = packages_dir

	p := tea.NewProgram(m)
	final_model, err := p.Run()
	if err != nil {
		fmt.Printf("Error running program: %v\n", err)
		return
	}

	finalModel := final_model.(model)
	if finalModel.step == 3 && finalModel.cursor == 0 {
		fmt.Println("\n" + titleStyle.Render("Creating Package"))
		indent := "    "
		sectionStyle := lipgloss.NewStyle().Bold(true)

		sanitized_name := sanitize_package_name(finalModel.packageName)
		intunewinFile := fmt.Sprintf("%s.intunewin", sanitized_name)
		installerFile := fmt.Sprintf("%s.%s",
			sanitized_name,
			strings.ToLower(finalModel.installerType))

		if finalModel.mode == "Repackage Application" {
			fmt.Println("\n" + sectionStyle.Render("Actions"))
			fmt.Printf("%s• Repackaging existing application...\n", indent)

			files, err := os.ReadDir(finalModel.outputDir)
			if err != nil {
				fmt.Printf("Error reading package directory: %v\n", err)
				return
			}

			var installer_path string
			for _, file := range files {
				if !file.IsDir() && (strings.HasSuffix(strings.ToLower(file.Name()), ".msi") ||
					strings.HasSuffix(strings.ToLower(file.Name()), ".exe")) {
					installer_path = filepath.Join(finalModel.outputDir, file.Name())
					installerFile = file.Name()
					break
				}
			}

			if installer_path == "" {
				fmt.Printf("Error: No installer file found in package directory\n")
				return
			}

			if strings.HasSuffix(strings.ToLower(installerFile), ".msi") {
				finalModel.installerType = "MSI"
			} else {
				finalModel.installerType = "EXE"
			}

			install_script_path := filepath.Join(finalModel.outputDir, "Install.ps1")
			uninstall_script_path := filepath.Join(finalModel.outputDir, "Uninstall.ps1")

			if _, err := os.Stat(install_script_path); err == nil {
				fmt.Printf("%s• Using existing Install.ps1 script\n", indent)
			} else {
				fmt.Printf("%s• Warning: Install.ps1 not found\n", indent)
			}

			if _, err := os.Stat(uninstall_script_path); err == nil {
				fmt.Printf("%s• Using existing Uninstall.ps1 script\n", indent)
			} else {
				fmt.Printf("%s• Warning: Uninstall.ps1 not found\n", indent)
			}

			fmt.Printf("%s• Generating IntuneWin package...\n", indent)
			args := []string{
				"-c", finalModel.outputDir,
				"-s", installer_path,
				"-o", finalModel.outputDir,
				"-q",
			}
			cmd := exec.Command(intuneUtilPath, args...)
			if output, err := cmd.CombinedOutput(); err != nil {
				fmt.Printf("Error generating IntuneWin package: %v\n%s\n", err, output)
				return
			}

			if finalModel.installerType == "MSI" {
				product_code, version, err := getMSIProductCode(installer_path)
				if err != nil {
					fmt.Printf("Warning: Could not get MSI product code: %v\n", err)
				} else {
					finalModel.productCode = product_code
					finalModel.version = version
				}
			}
		} else if finalModel.source == "Local File" {
			fmt.Println("\n" + sectionStyle.Render("Actions"))
			fmt.Printf("%s• Copying installer file...\n", indent)
			if err := copyFileToDir(finalModel.textInput, finalModel.outputDir, installerFile); err != nil {
				fmt.Printf("Error copying installer: %v\n", err)
				return
			}

			if finalModel.installerType == "MSI" {
				msiPath := filepath.Join(finalModel.outputDir, installerFile)
				productCode, version, err := getMSIProductCode(msiPath)
				if err != nil {
					fmt.Printf("Error getting product code: %v\n", err)
					return
				}
				finalModel.productCode = productCode
				finalModel.version = version
			}

			if err := createPackageScripts(finalModel.outputDir, finalModel.packageName, finalModel.installerType); err != nil {
				fmt.Printf("Error creating package scripts: %v\n", err)
				return
			}

			fmt.Printf("%s• Generating IntuneWin package...\n", indent)
			args := []string{
				"-c", finalModel.outputDir,
				"-s", filepath.Join(finalModel.outputDir, installerFile),
				"-o", finalModel.outputDir,
				"-q",
			}
			cmd := exec.Command(intuneUtilPath, args...)
			if output, err := cmd.CombinedOutput(); err != nil {
				fmt.Printf("Error generating IntuneWin package: %v\n%s\n", err, output)
				return
			}

		} else {
			fmt.Println("\n" + sectionStyle.Render("Actions"))
			fmt.Printf("%s• Downloading installer file...\n", indent)

			download_path := filepath.Join(downloadsDir, installerFile)

			if err := downloadFile(finalModel.textInput, download_path); err != nil {
				fmt.Printf("Error downloading installer: %v\n", err)
				return
			}

			fmt.Printf("%s• Copying installer to package directory...\n", indent)
			if err := copyFileToDir(download_path, finalModel.outputDir, installerFile); err != nil {
				fmt.Printf("Error copying installer: %v\n", err)
				return
			}

			if finalModel.installerType == "MSI" {
				msiPath := filepath.Join(finalModel.outputDir, installerFile)
				productCode, version, err := getMSIProductCode(msiPath)
				if err != nil {
					fmt.Printf("Error getting product code: %v\n", err)
					return
				}
				finalModel.productCode = productCode
				finalModel.version = version
			}

			if err := createPackageScripts(finalModel.outputDir, finalModel.packageName, finalModel.installerType); err != nil {
				fmt.Printf("Error creating package scripts: %v\n", err)
				return
			}

			fmt.Printf("%s• Generating IntuneWin package...\n", indent)
			args := []string{
				"-c", finalModel.outputDir,
				"-s", filepath.Join(finalModel.outputDir, installerFile),
				"-o", finalModel.outputDir,
				"-q",
			}
			cmd := exec.Command(intuneUtilPath, args...)
			if output, err := cmd.CombinedOutput(); err != nil {
				fmt.Printf("Error generating IntuneWin package: %v\n%s\n", err, output)
				return
			}
		}

		fmt.Println("\n" + titleStyle.Render("Package Complete"))

		fmt.Println("\n" + sectionStyle.Render("Package Summary"))
		fmt.Printf("%s• Name: %s\n", indent, finalModel.packageName)
		fmt.Printf("%s• Type: %s\n", indent, finalModel.installerType)
		if finalModel.version != "" {
			fmt.Printf("%s• Version: %s\n", indent, finalModel.version)
		}
		if finalModel.installerType == "MSI" {
			fmt.Printf("%s• Product Code: %s\n", indent, finalModel.productCode)
		}

		fmt.Println("\n" + sectionStyle.Render("Package Location"))
		fmt.Printf("%s• Source: %s\n", indent, finalModel.textInput)
		fmt.Printf("%s• Package Directory: %s\n", indent, finalModel.outputDir)
		fmt.Printf("%s• Installer File: %s\n", indent, installerFile)
		fmt.Printf("%s• IntuneWin File: %s\n", indent, intunewinFile)

		fmt.Println("\n" + sectionStyle.Render("Intune Configuration"))

		fmt.Printf("%s• Install Script: Install.ps1\n", indent)
		fmt.Printf("%s• Uninstall Script: Uninstall.ps1\n", indent)

		fmt.Println("\n" + sectionStyle.Render("Detection Method"))
		if finalModel.installerType == "MSI" {
			fmt.Printf("%s• MSI Product Code:\n", indent)
			fmt.Printf("%s  - Property: ProductCode\n", indent)
			fmt.Printf("%s  - Value: %s\n", indent, finalModel.productCode)
			if finalModel.version != "" {
				fmt.Printf("%s• Version Detection:\n", indent)
				fmt.Printf("%s  - Property: ProductVersion\n", indent)
				fmt.Printf("%s  - Value: %s\n", indent, finalModel.version)
				fmt.Printf("%s  - Operator: Greater than or equal to\n", indent)
			}
		} else {
			fmt.Printf("%s• Use custom detection script or file existence\n", indent)
		}

		fmt.Println("\n" + sectionStyle.Render("Additional Configuration"))
		fmt.Printf("%s• To modify installation arguments:\n", indent)
		if finalModel.installerType == "MSI" {
			fmt.Printf("%s  - Default: /qn /norestart\n", indent)
			fmt.Printf("%s  - Edit Install.ps1 and modify $install_args\n", indent)
		} else {
			fmt.Printf("%s  - Default: /silent\n", indent)
			fmt.Printf("%s  - Edit Install.ps1 and modify $install_args\n", indent)
		}
		fmt.Printf("%s• To add custom installation steps:\n", indent)
		fmt.Printf("%s  - Edit Install.ps1 in the package directory\n", indent)
		fmt.Printf("%s  - Add steps before or after the install_application function\n", indent)
		fmt.Println()
	}
}

func ensureNexusDirs() error {
	dirs := []string{nexusDir, intuneUtilDir, packagesDir, downloadsDir}
	for _, dir := range dirs {
		if err := os.MkdirAll(dir, 0755); err != nil {
			return fmt.Errorf("failed to create directory %s: %v", dir, err)
		}
	}
	return nil
}

func downloadFile(url, filepath string) error {
	dir := path.Dir(filepath)
	fmt.Printf("    Creating directory: %s\n", dir)

	if err := os.MkdirAll(dir, 0755); err != nil {
		return fmt.Errorf("failed to create directory %s: %v", dir, err)
	}

	if _, err := os.Stat(dir); os.IsNotExist(err) {
		return fmt.Errorf("directory creation failed, path still doesn't exist: %s", dir)
	}

	fmt.Printf("    Opening file for writing: %s\n", filepath)
	out, err := os.Create(filepath)
	if err != nil {
		return fmt.Errorf("failed to create file: %v", err)
	}
	defer out.Close()

	resp, err := http.Get(url)
	if err != nil {
		return fmt.Errorf("failed to download file: %v", err)
	}
	defer resp.Body.Close()

	if resp.StatusCode != http.StatusOK {
		return fmt.Errorf("bad status: %s", resp.Status)
	}

	size := resp.ContentLength
	fmt.Printf("    Downloading %s (%d bytes)...\n", path.Base(filepath), size)

	_, err = io.Copy(out, resp.Body)
	if err != nil {
		return fmt.Errorf("failed to save downloaded file: %v", err)
	}

	return nil
}

func ensureIntuneUtil() error {
	if _, err := os.Stat(intuneUtilPath); err == nil {
		return nil
	}

	fmt.Println("Downloading IntuneWinAppUtil.exe...")
	if err := downloadFile(intuneUtilUrl, intuneUtilPath); err != nil {
		return fmt.Errorf("failed to download IntuneWinAppUtil: %v", err)
	}

	return nil
}

func sanitizePackageName(name string) string {
	name = strings.ToLower(name)
	name = strings.ReplaceAll(name, " ", "-")

	reg := regexp.MustCompile(`[^a-z0-9\-]`)
	name = reg.ReplaceAllString(name, "")

	reg = regexp.MustCompile(`-+`)
	name = reg.ReplaceAllString(name, "-")

	return strings.Trim(name, "-")
}

func get_existing_packages(packages_dir string) ([]string, error) {
	if err := os.MkdirAll(packages_dir, 0755); err != nil {
		return nil, fmt.Errorf("failed to create packages directory: %v", err)
	}

	entries, err := os.ReadDir(packages_dir)
	if err != nil {
		return nil, fmt.Errorf("failed to read packages directory: %v", err)
	}

	var packages []string
	caser := cases.Title(language.English)
	for _, entry := range entries {
		if entry.IsDir() {
			name := entry.Name()
			name = strings.ReplaceAll(name, "-", " ")
			name = caser.String(name)
			packages = append(packages, name)
		}
	}

	return packages, nil
}

func sanitize_package_name(name string) string {
	return sanitizePackageName(name)
}

func load_config() (string, error) {
	config_path := filepath.Join(nexusDir, "config.json")
	if _, err := os.Stat(config_path); os.IsNotExist(err) {
		return packagesDir, nil
	}

	data, err := os.ReadFile(config_path)
	if err != nil {
		return packagesDir, err
	}

	var config struct {
		PackagesDir string `json:"packages_dir"`
	}

	if err := json.Unmarshal(data, &config); err != nil {
		return packagesDir, err
	}

	if config.PackagesDir == "" {
		return packagesDir, nil
	}

	return config.PackagesDir, nil
}

func save_config(packages_dir string) error {
	config_path := filepath.Join(nexusDir, "config.json")

	config := struct {
		PackagesDir string `json:"packages_dir"`
	}{
		PackagesDir: packages_dir,
	}

	data, err := json.Marshal(config)
	if err != nil {
		return err
	}

	return os.WriteFile(config_path, data, 0644)
}
