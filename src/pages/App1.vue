<template>
  <q-card class="card q-ma-md bg-blue flex justify-center text-white">
    <q-card-section>
      <q-card class="bg-white text-grey-8">
        <q-card-section>
          <span style="font-size: 18px; font-weight: 700">App de Gerenciamento</span>
        </q-card-section>
        <q-card-section>
          <q-file
            v-model="xslxFile"
            label="Selecione o arquivo .xlsx"
            filled
            style="min-width: 300px"
          >
            <template v-slot:after v-if="xslxFile">
              <q-btn
                color="primary"
                dense
                icon="cloud_upload"
                round
                @click="onFileChange"
                :disable="!xslxFile"
              />
            </template>
          </q-file>
        </q-card-section>
        <q-card-section v-if="showTable">
          <q-btn v-if="jsonFromServeFile" class="q-mb-sm" icon="download" @click="downloadXLSX(jsonFromServeFile)" label="Baixar .xlsx gerado"></q-btn>
          <q-btn v-else @click="runSeverUpload" class="q-mb-sm" icon="play" label="Executar"></q-btn>
          <q-table
            style="max-width: 800px"
            :rows="rows"
            title="Formato tabela"
            :columns="columns"
            row-key="name"
          />
          <q-table
            v-if="showTable"
            style="max-width: 800px"
            :rows="rows"
            grid
            class="q-mt-md"
            title="Formato cards"
            :columns="columns"
            row-key="name"
          />
          <q-field label="Formato JSON" style="max-width: 800px" class="q-mt-md" v-if="showTable" stack-label>
            <template v-slot:control>
              <div class="self-center no-outline" tabindex="0">{{ (xslxJson) }}</div>
            </template>
          </q-field>
        </q-card-section>
      </q-card>
    </q-card-section>
  </q-card>
</template>

<script>
import readXlsxFile from 'read-excel-file'
import axios from "axios";
const XLSX = require('xlsx');

export default {
  name: "App1",
  methods: {
    convertTabletoJson(data){
      const arr = [];
      for(let i = 1; i < data.length; i++){
        arr.push({
          id: data[0][0] === 'LOTE' ? data[i][0] : null,
          yellowSubBatch: {
            id: data[0][4] === 'AMARELO' ? data[i][4] : null,
            avgWeight: data[0][6] === 'PESO MEDIO AMARELO' ? data[i][6] : null,
            type: 1,
            quantity: data[0][5] === 'QUANTIDADE AMARELO' ? data[i][5] : null},
          greenSubBatch: {
            id: data[0][1] === 'VERDE' ? data[i][1] : null,
            avgWeight: data[0][3] === 'PESO MEDIO VERDE' ? data[i][3] : null,
            type: 1,
            quantity: data[0][2] === 'QUANTIDADE VERDE' ? data[i][2] : null}}
        )}
      this.xslxJson = arr;
    },

    convertJsonToTable(data){
      const arr = [];
      for(let i = 0; i < data.length; i++){
        arr.push({
          "LOTE": data[i].id,
          "AMARELO": data[i].yellowSubBatch.id,
          "QUANTIDADE AMARELO": data[i].yellowSubBatch.quantity,
          "PESO MEDIO AMARELO": data[i].yellowSubBatch.avgWeight,
          "VERDE": data[i].greenSubBatch.id,
          "QUANTIDADE VERDE": data[i].greenSubBatch.quantity,
          "PESO MEDIO VERDE": data[i].greenSubBatch.avgWeight,
          "DIFERENCA": data[i].quantityDiff,
          "DIFERENCA PESO": data[i].weightDiff
        })
      }
      console.log(arr)
      return arr;
    },

    downloadXLSX(file){
      const workSheet = XLSX.utils.json_to_sheet(this.convertJsonToTable(file));
      const workBook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workBook, workSheet, 'Tabela Gerada');
      XLSX.writeFile(workBook, 'tabela_gerada.xlsx');
    },

    createTable(rows){
      let rowObj = {};

      rows[0].forEach((item) =>{
        this.columns.push({name: item, required: true, label: item, align: 'left', field: item,  sortable: true})
      })

      for(let i = 1; i < rows.length; i++){
        rows[i].forEach((item, idx) =>{
          rowObj[this.columns[idx].name] = item;
        })
        this.rows.push(rowObj)
        rowObj = {};
      }
      this.showTable = true;
    },

    runSeverUpload(){
      return new Promise(
        () => {
          this.$axios.post("https://cors-anywhere.herokuapp.com/" + "https://4kpw53thkri4yk3p4lcdclcnxq0achir.lambda-url.us-east-1.on.aws/gripenew/optimize", this.xslxJson, {
            headers: {
              "access-control-allow-origin": "*",
              "content-type": "application/json; charset=utf-8",
              "x-amzn-requestid": "4cd1822c-1c8b-4530-8df5-9b7a6f5eed33",
              "x-amzn-trace-id": "root=1-634f49e4-3c9877944d04a1553205db4f;sampled=0",
              "x-final-url": "https://4kpw53thkri4yk3p4lcdclcnxq0achir.lambda-url.us-east-1.on.aws/gripenew/optimize"

            }, }).then((result) =>{
              this.jsonFromServeFile = result.data;
              console.log(this.jsonFromServeFile);
          })
        });
    },

    onFileChange() {
      readXlsxFile(this.xslxFile).then((rows) => {
        this.createTable(rows);
        this.convertTabletoJson(rows);
      });
    }
  },

  data(){
    const xslxFile = null;
    const columns = [];
    const showTable = false;
    const rows = [];
    const xslxJson = null;
    const jsonFromServeFile = null;

    const outputJson = []

    return{
      xslxFile: xslxFile,
      columns: columns,
      rows: rows,
      showTable: showTable,
      xslxJson: xslxJson,
      jsonFromServeFile: jsonFromServeFile,
      axiosInstace: axios.create({
        baseURL: 'https://4kpw53thkri4yk3p4lcdclcnxq0achir.lambda-url.us-east-1.on.aws/',
        headers: {"Access-Control-Allow-Origin": "*"}
      })
    }
  }
}
</script>

<style scoped>

</style>
