const express = require('express');
const path = require('path');
const app = express();
// Serve wwwroot (or dist) with directory index support
app.use(express.static(path.join(__dirname, 'wwwroot')));
const port = process.env.PORT || 8080;
app.listen(port, () => console.log(`Listening on ${port}`));