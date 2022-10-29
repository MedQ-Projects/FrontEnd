<template>
  <q-page>
    <q-card class="q-ma-md justify-center bg-blue flex text-white">
      <q-card :key="key" class=" q-ma-md flex column bg-white text-grey-8">
        <q-card-section>
          <span style="font-size: 18px; font-weight: 700">App de Gerenciamento</span>
          <div class="shadow-2 q-pa-sm q-mt-sm">
            <span>Descrição da aplicação: <br> Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.
            </span>
          </div>
        </q-card-section>
        <q-card-section>
          <q-file
            :disable="loading"
            @update:model-value="updateFiles"
            v-model="xslxFile"
            label="Selecione o arquivo .xlsx"
            filled
            color="secondary"
            clearable
            accept=".xlsx"
          >
            <template v-slot:after v-if="canUpload && !xslxJson">
              <q-btn
                color="primary"
                dense
                icon="cloud_upload"
                round
                @click="onFileChange"
              />
            </template>
          </q-file>
        </q-card-section>
        <q-card-section v-if="showTable">
          <q-btn v-if="jsonFromServeFile" class="q-mb-sm" color="secondary" icon="download" @click="downloadXLSX(jsonFromServeFile)" label="Baixar .xlsx gerado"></q-btn>
          <q-btn v-else @click="runSeverUpload" color="primary" class="q-mb-sm" icon="play_arrow" label="Executar"></q-btn>
          <q-linear-progress size="15px"  :value="progress" color="positive" class="q-mt-sm q-mb-sm" />
          <q-table
            :rows="rows"
            :loading="loading"
            title="Formato tabela"
            :columns="columns"
            row-key="name"
          >
            <template v-slot:loading>
              <q-inner-loading showing color="primary" />
            </template>
          </q-table>
          <q-table
            v-if="showTable"
            :rows="rows"
            grid
            :loading="loading"
            class="q-mt-md"
            title="Formato cards"
            :columns="columns"
            row-key="name"
          >
            <template v-slot:loading>
              <q-inner-loading showing color="primary" />
            </template>
          </q-table>
        </q-card-section>
      </q-card>
    </q-card>
  </q-page>
</template>

<script>
import readXlsxFile from 'read-excel-file'
import {useQuasar} from "quasar";
const XLSX = require('xlsx');
const $q = useQuasar()

const headers = {
  'Content-Type': 'application/json',
  'Access-Control-Allow-Origin':  '*'
}

export default {
  name: "App1",
  computed: {
    canUpload(){
      return this.xslxFile !== null;
    }
  },
  methods: {
    updateFiles(){
      this.xslxFile = null;
      this.columns = [];
      this.showTable = false;
      this.rows = [];
      this.xslxJson = null;
      this.jsonFromServeFile = null;
      this.loading = false;
      this.progress = 0;
    },

    convertXLSXtoJson(data){
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

    convertJsonToXLSX(data){
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
      return arr;
    },

    downloadXLSX(file){
      const workSheet = XLSX.utils.json_to_sheet(this.convertJsonToXLSX(file));
      const workBook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workBook, workSheet, 'Tabela Gerada');
      XLSX.writeFile(workBook, 'tabela_gerada.xlsx');
    },

    createTableFromXLSX(rows){
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

    createTableFromJson(jsonFile){
      this.columns = [];
      this.rows = [];
      const fields = ['LOTE', 'AMARELO', 'QUANTIDADE AMARELO', 'PESO MEDIO AMARELO', 'VERDE', 'QUANTIDADE VERDE', 'PESO MEDIO VERDE', 'DIFERENCA', 'DIFERENCA PESO']
      fields.forEach(field => {
        this.columns.push({name: field, required: true, label: field, align: 'left', field: field,  sortable: true})
      })
      jsonFile.forEach((item) =>{
        this.rows.push({
          'LOTE': item.id,
          'AMARELO': item.yellowSubBatch.id,
          'QUANTIDADE AMARELO': item.yellowSubBatch.quantity,
          'PESO MEDIO AMARELO': item.yellowSubBatch.avgWeight,
          'VERDE': item.greenSubBatch.id,
          'QUANTIDADE VERDE': item.greenSubBatch.quantity,
          'PESO MEDIO VERDE': item.greenSubBatch.avgWeight,
          'DIFERENCA': item.quantityDiff,
          'DIFERENCA PESO': item.weightDiff
          })
      })
      this.key = this.key + 1;
    },

    runSeverUpload(){
      this.loading = true;
      return new Promise(
        () => {
          this.$axios.post("https://4kpw53thkri4yk3p4lcdclcnxq0achir.lambda-url.us-east-1.on.aws/gripenew/optimize", this.xslxJson, {
            headers: headers
          })
            .then((result) =>{
              this.jsonFromServeFile = result.data;
              this.createTableFromJson(this.jsonFromServeFile);
              this.loading = false;

              this.progress = 1;
          }).catch(err =>{
            if(err){
              this.loading = false;
              this.$q.notify({
                message: 'Erro na Execução !',
                icon: 'error',
                caption: 'Verifique se o arquivo .xlsx está correto.',
                color: 'negative'
              })
            }
            console.log(err)
          })
        });
    },

    onFileChange() {
      readXlsxFile(this.xslxFile).then((rows) => {
        this.createTableFromXLSX(rows);
        this.convertXLSXtoJson(rows);
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
    const loading = false;

    return{
      xslxFile: xslxFile,
      columns: columns,
      rows: rows,
      showTable: showTable,
      xslxJson: xslxJson,
      jsonFromServeFile: jsonFromServeFile,
      loading: loading,
      progress: 0,
      key: 0
    }
  }
}
</script>

<style scoped>

</style>
