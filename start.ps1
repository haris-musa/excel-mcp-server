# Excel MCP Server Windows PowerShell Start Script

param(
    [string]$Action = "start"
)

# Color definitions
$Colors = @{
    Red = 'Red'
    Green = 'Green'
    Yellow = 'Yellow'
    Blue = 'Blue'
    Default = 'White'
}

# Logging functions
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "Info"
    )
    
    switch ($Level) {
        "Info" { Write-Host "[INFO] $Message" -ForegroundColor $Colors.Green }
        "Warn" { Write-Host "[WARN] $Message" -ForegroundColor $Colors.Yellow }
        "Error" { Write-Host "[ERROR] $Message" -ForegroundColor $Colors.Red }
        default { Write-Host "[INFO] $Message" -ForegroundColor $Colors.Default }
    }
}

# Check if Docker is installed
function Test-Docker {
    if (-not (Get-Command docker -ErrorAction SilentlyContinue)) {
        Write-Log "Docker is not installed. Please install Docker Desktop first." -Level "Error"
        exit 1
    }
    
    if (-not (Get-Command docker-compose -ErrorAction SilentlyContinue)) {
        Write-Log "Docker Compose is not installed. Please install Docker Compose first." -Level "Error"
        exit 1
    }
}

# Create necessary directories
function New-Directories {
    Write-Log "Creating necessary directories..."
    
    # Read .env file
    $envFile = ".\.env"
    $hostExcelPath = ".\excel_files"
    $hostLogPath = ".\logs"
    
    if (Test-Path $envFile) {
        $envContent = Get-Content $envFile
        foreach ($line in $envContent) {
            if ($line -match "HOST_EXCEL_PATH=(.+)") {
                $hostExcelPath = $matches[1]
            }
            if ($line -match "HOST_LOG_PATH=(.+)") {
                $hostLogPath = $matches[1]
            }
        }
    }
    
    # Create directories
    if (-not (Test-Path $hostExcelPath)) {
        New-Item -ItemType Directory -Path $hostExcelPath -Force | Out-Null
    }
    
    if (-not (Test-Path $hostLogPath)) {
        New-Item -ItemType Directory -Path $hostLogPath -Force | Out-Null
    }
    
    Write-Log "Excel files directory: $hostExcelPath"
    Write-Log "Logs directory: $hostLogPath"
}

# Check environment file
function Test-EnvFile {
    if (-not (Test-Path ".\.env")) {
        Write-Log ".env file does not exist, using default configuration" -Level "Warn"
        Write-Log "You can copy env.example to create .env file"
        
        # Create default .env file
        $defaultEnv = @"
# Excel MCP Server Configuration
HOST_EXCEL_PATH=./excel_files
HOST_LOG_PATH=./logs
HOST_PORT=8000
FASTMCP_PORT=8000
"@
        $defaultEnv | Out-File -FilePath ".\.env" -Encoding UTF8
        Write-Log "Default .env file created"
    }
}

# Build and start service
function Start-Service {
    Write-Log "Building and starting Excel MCP Server..."
    
    # Build image
    docker-compose build
    
    if ($LASTEXITCODE -ne 0) {
        Write-Log "Build failed!" -Level "Error"
        exit 1
    }
    
    # Start service
    docker-compose up -d
    
    if ($LASTEXITCODE -ne 0) {
        Write-Log "Start failed!" -Level "Error"
        exit 1
    }
    
    # Wait for service to start
    Write-Log "Waiting for service to start..."
    Start-Sleep -Seconds 5
    
    # Check service status
    $status = docker-compose ps
    if ($status -match "Up") {
        Write-Log "Service started successfully!"
        
        # Read port configuration
        $hostPort = "8000"
        $hostExcelPath = "./excel_files"
        $hostLogPath = "./logs"
        
        if (Test-Path ".\.env") {
            $envContent = Get-Content ".\.env"
            foreach ($line in $envContent) {
                if ($line -match "HOST_PORT=(.+)") {
                    $hostPort = $matches[1]
                }
                if ($line -match "HOST_EXCEL_PATH=(.+)") {
                    $hostExcelPath = $matches[1]
                }
                if ($line -match "HOST_LOG_PATH=(.+)") {
                    $hostLogPath = $matches[1]
                }
            }
        }
        
        # Show access information
        Write-Host ""
        Write-Log "Service access information:"
        Write-Log "  MCP endpoint: http://localhost:$hostPort/mcp"
        Write-Log "  Excel files directory: $hostExcelPath"
        Write-Log "  Logs directory: $hostLogPath"
        Write-Host ""
        
        # Show client configuration
        Write-Log "MCP client configuration:"
        Write-Host @"
{
   "mcpServers": {
      "excel": {
         "url": "http://localhost:$hostPort/mcp"
      }
   }
}
"@ -ForegroundColor $Colors.Blue
        
    } else {
        Write-Log "Service start failed!" -Level "Error"
        Write-Log "Please check logs: docker-compose logs excel-mcp-server" -Level "Error"
        exit 1
    }
}

# Show help information
function Show-Help {
    Write-Host "Excel MCP Server Windows PowerShell Start Script"
    Write-Host ""
    Write-Host "Usage: .\start.ps1 [option]"
    Write-Host ""
    Write-Host "Options:"
    Write-Host "  start     Start service (default)"
    Write-Host "  stop      Stop service"
    Write-Host "  restart   Restart service"
    Write-Host "  status    View service status"
    Write-Host "  logs      View service logs"
    Write-Host "  help      Show this help"
    Write-Host ""
}

# Main function
function Main {
    switch ($Action.ToLower()) {
        "start" {
            Test-Docker
            Test-EnvFile
            New-Directories
            Start-Service
        }
        "stop" {
            Write-Log "Stopping service..."
            docker-compose down
            Write-Log "Service stopped"
        }
        "restart" {
            Write-Log "Restarting service..."
            docker-compose down
            Start-Sleep -Seconds 2
            docker-compose up -d
            Write-Log "Service restarted"
        }
        "status" {
            Write-Log "Service status:"
            docker-compose ps
        }
        "logs" {
            Write-Log "Service logs:"
            docker-compose logs -f excel-mcp-server
        }
        "help" {
            Show-Help
        }
        default {
            Write-Log "Unknown option: $Action" -Level "Error"
            Show-Help
            exit 1
        }
    }
}

# Execute main function
Main 