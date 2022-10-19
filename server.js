const
  express = require('express'),
  cors = require('cors'),
  serveStatic = require('serve-static'),
  history = require('connect-history-api-fallback'),
  port = process.env.PORT || 5000

const app = express()

app.use(cors())
app.use(history())
app.use(serveStatic(__dirname + '/dist/spa'))
app.listen(port)
