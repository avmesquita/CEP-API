const express = require('express')
const bodyParser = require('body-parser')
const port = process.env.PORT || 10000

const path = require('path')
const knex = require('knex')({
  client: 'sqlite3',
  connection: { filename: './cep.db3' },
  useNullAsDefault: true,
})

const app = express()

app.use(bodyParser.urlencoded({ extended: false }))

app.get('/api/v1/cep/:filter', (req, res) => {

 filter = req.params.filter

 if (filter != '') {
     list = () =>  knex('cep').select('*')
	                   .orWhere('txt_cep', filter)	 
	                   .orWhere('txt_cidade_uf', 'like', '%'+filter+'%')	 
					   .orWhere('txt_bairro', 'like', '%'+filter+'%')	 
					   .orWhere('txt_localidade', 'like', '%'+filter+'%')
					   .limit(10)
 } else {
	 list = () =>  knex('cep').select('*').limit(10)	 
 }
  list()
    .then(data => res.json(data))
})

app.get('/', (req, res) => {  

  commandDefault = "http://" + req.headers.host + "/api/v1/cep/<value>"

  defaultMessage = JSON.stringify({ 'usage' : commandDefault });

  res.json(defaultMessage)
})

app.listen(port, () => {
  console.log(`\nCEP REST API started on http://localhost:${port}\n\nPress Ctrl+C to terminate.`)
})
