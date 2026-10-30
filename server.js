const express = require('express');
const fs = require('fs');
const path = require('path');
const docx = require('docx'); // Run: npm install docx express

const app = express();
app.use(express.urlencoded({ extended: true }));
app.use(express.static('public'));

app.post('/submit-order', async (req, res) => {
    const { name, email, origamiType, quantity, notes } = req.body;
    const orderDate = new Date().toLocaleString();

    // Create the Word document structure
    const doc = new docx.Document({
        sections: [{
            properties: {},
            children: [
                new docx.Paragraph({
                    children: [
                        new docx.TextRun({ text: "THE ORIGIN OF CRAFT", bold: true, size: 28, color: "D4A373" }),
                    ],
                }),
                new docx.Paragraph({ text: "New Customer Order Invoice", italics: true, spacing: { after: 400 } }),
                
                new docx.Paragraph({ children: [ new docx.TextRun({ text: "Order Details:", bold: true }) ] }),
                new docx.Paragraph({ text: `Date: ${orderDate}` }),
                new docx.Paragraph({ text: `Customer Name: ${name}` }),
                new docx.Paragraph({ text: `Email: ${email}` }),
                new docx.Paragraph({ text: `Product Ordered: ${origamiType}` }),
                new docx.Paragraph({ text: `Quantity: ${quantity}` }),
                new docx.Paragraph({ text: `Special Instructions: ${notes || 'None'}` }),
            ],
        }],
    });

    // Write the compiled document to a file named after the customer
    const fileName = `Order_${name.replace(/\s+/g, '_')}_${Date.now()}.docx`;
    const filePath = path.join(__dirname, 'orders', fileName);

    // Make sure the orders folder exists
    if (!fs.existsSync(path.join(__dirname, 'orders'))){
        fs.mkdirSync(path.join(__dirname, 'orders'));
    }

    const b64string = await docx.Packer.toBase64String(doc);
    fs.writeFileSync(filePath, Buffer.from(b64string, 'base64'));

    // Respond back to the user website
    res.send('<h1>Thank you for your order! The Origin of Craft team has received it.</h1><a href="/">Go Back</a>');
});

app.listen(3000, () => console.log('Server running on http://localhost:3000'));
