import { renderToString } from "preact-render-to-string";

export function generatePdfFromSpreadsheetData(data: ReportProps): GoogleAppsScript.HTML.HtmlOutput {
    // 1. Render TSX component directly to an HTML string at runtime
    const htmlString = renderToString(ReportTemplate(data));

    // 2. Convert HTML string to PDF Blob
    return HtmlService.createHtmlOutput(htmlString);

}

interface LineItem {
    description: string;
    quantity: number;
    unitPrice: number;
}

interface ReportProps {
    title: string;
    clientName: string;
    items: LineItem[];
    notes?: string;
}

export function ReportTemplate({ title, clientName, items, notes }: ReportProps) {
    const total = items.reduce((acc, item) => acc + item.quantity * item.unitPrice, 0);

    return (
        <html>
            <head>
                <style>{`
          @page { size: A4; margin: 20mm; }
          body { font-family: Arial, sans-serif; color: #333; line-height: 1.4; }
          .header { display: flex; justify-content: space-between; align-items: center; }
          .logo { height: 50px; }
          table { width: 100%; border-collapse: collapse; margin-top: 20px; }
          th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
          th { background-color: #f4f4f4; }
          .total { text-align: right; font-weight: bold; margin-top: 15px; }
        `}</style>
            </head>
            <body>
                <div class="header">
                    <h1>{title}</h1>
                    {/* <img src={logoUrl} class="logo" alt="Logo" /> */}
                </div>

                <p>
                    <strong>Prepared for:</strong> {clientName}
                </p>

                {/* Dynamic Table: Rows adapt cleanly depending on input data */}
                <table>
                    <thead>
                        <tr>
                            <th>Description</th>
                            <th>Qty</th>
                            <th>Price</th>
                            <th>Subtotal</th>
                        </tr>
                    </thead>
                    <tbody>
                        {items.map((item, idx) => (
                            <tr key={idx}>
                                <td>{item.description}</td>
                                <td>{item.quantity}</td>
                                <td>${item.unitPrice.toFixed(2)}</td>
                                <td>${(item.quantity * item.unitPrice).toFixed(2)}</td>
                            </tr>
                        ))}
                    </tbody>
                </table>

                <div class="total">Total Amount: ${total.toFixed(2)}</div>

                {/* Conditionally rendered section */}
                {notes && (
                    <div style={{ marginTop: "30px", padding: "10px", background: "#f9f9f9" }}>
                        <strong>Notes:</strong> {notes}
                    </div>
                )}
            </body>
        </html>
    );
}
