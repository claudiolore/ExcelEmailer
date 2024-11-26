# Excel Email Sender

A C# console application that automates sending personalized emails with PDF attachments using data from Excel spreadsheets.

## Features

- Read recipient data (name, address, email) from Excel files
- Personalize PDF attachments for each recipient
- Support for Gmail SMTP server
- Interactive console interface with color-coded messages
- Detailed error reporting and success tracking
- Batch email processing with progress tracking

## Prerequisites

- .NET Framework (version compatible with EPPlus)
- Required NuGet Packages:
  - EPPlus (for Excel file handling)
  - MailKit (for email functionality)
  - MimeKit (for email message construction)
  - iTextSharp (for PDF manipulation)

## Configuration

### Excel File Format
The Excel file should contain the following columns:
- Nominativo (Name)
- Indirizzo (Address)
- Email

### SMTP Settings
The application is configured to use Gmail SMTP server with the following settings:
- Server: smtp.gmail.it
- Port: 587
- Security: SSL/TLS

## Usage

1. Launch the application
2. Follow the interactive prompts:
   - Enter the path to your Excel file
   - Provide email subject and body
   - Choose whether to attach a PDF
   - Enter sender email credentials
3. Review the summary before sending
4. Monitor the sending progress with real-time status updates

## Error Handling

The application includes comprehensive error handling for:
- Invalid file paths
- Missing Excel columns
- SMTP connection issues
- Email sending failures
- Invalid credentials

## Output

The application provides:
- Real-time sending status for each email
- Color-coded console messages for better visibility
- Final report showing successful and failed email counts

## Security Notes

- Email credentials are handled securely in memory
- SSL certificate validation can be configured for production use
- The application supports timeout configuration for SMTP operations

## Contributing

1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Push to the branch
5. Create a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.
