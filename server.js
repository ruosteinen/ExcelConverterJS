const express = require('express');
const app = express();
const host = '192.168.1.8';
const port = 8080;

app.use(express.static('sankey'));

app.listen(port, host, () => {
  console.log(`Server running at http://${host}:${port}/`);
});
