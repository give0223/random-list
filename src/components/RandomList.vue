<template>
  <div class="hello">
    <el-card class="card-box" shadow="always" :body-style="{ padding: '20px' }">
      <div slot="header">
        <h1>{{ msg }}</h1>
      </div>
      <div>
        <h2>當前抽驗文件: {{ targetData.currentFile }}</h2>
        <h3>當前抽驗文件內容總筆數: {{ targetData.maxNum }}</h3>
      </div>
    </el-card>
    <el-card class="card-box" shadow="always" :body-style="{ padding: '20px' }">
      <div class="target-bar">
        <el-upload
          class="upload-box"
          action="#"
          :auto-upload="true"
          :before-upload="handleUpload"
          :limit="1"
          accept=".xlsx,.xls"
          :show-file-list="false"
        >
          <el-button class="upload-btn" size="medium" type="primary"
            ><i class="el-icon-upload el-icon--right"
              >點擊上傳抽驗文件</i
            ></el-button
          >
          <div slot="tip" class="el-upload__tip textType">
            只能上傳xlsx/xls文件
          </div>
        </el-upload>
        <div class="target-set">
          <p>抽驗數量:</p>
          <el-input-number
            v-model="targetData.targetNum"
            placeholder="請輸入抽驗數量"
            size="normal"
            :max="targetData.maxNum"
            :min="1"
            @change="handleChange"
          ></el-input-number>
          <el-button
            class="random-btn"
            :disabled="targetData.workSheetArr.length <= 0"
            size="default"
            @click="handleRandom"
            >點擊抽驗</el-button
          >
        </div>
      </div>
    </el-card>
    <el-card
      v-if="targetData.targetList.length > 0"
      class="card-box"
      shadow="always"
      :body-style="{ padding: '20px' }"
    >
      <div slot="header">
        <span class="title">抽驗名單</span>
      </div>
      <div class="current-table">
        <el-table
          :data="targetData.tableData"
          element-loading-text="Loading..."
          empty-text="查無資料"
          v-loading="loading"
          style="width: 100%"
          :max-height="targetData.tableData.length > 0 ? '750' : '200'"
          :header-cell-style="{ background: '#2c3e50', color: '#ffffff' }"
        >
          <el-table-column type="index" label="序列" width="70" align="left">
          </el-table-column>
          <el-table-column prop="serialNum" label="流水號" width="120">
          </el-table-column>
          <el-table-column prop="productName" label="產品名稱" min-width="200">
          </el-table-column>
          <el-table-column prop="companyName" label="單位名稱" min-width="180">
          </el-table-column>
          <el-table-column prop="name" label="申請者姓名" width="120">
          </el-table-column>
          <el-table-column prop="origin" label="產地" min-width="180">
          </el-table-column>
          <el-table-column prop="area" label="申請地區" min-width="130">
          </el-table-column>
          <el-table-column prop="applicationNum" label="申請編號" width="150">
          </el-table-column>
          <el-table-column prop="volume" label="可出貨材積" min-width="120">
          </el-table-column>
          <el-table-column prop="state" label="狀態" width="100">
          </el-table-column>
        </el-table>
      </div>
    </el-card>
  </div>
</template>

<script>
import { read, utils, writeFile } from 'xlsx';
export default {
  name: 'RandomList',
  props: {
    msg: String,
  },
  data() {
    return {
      loading: false,
      targetData: {
        currentFile: '',
        maxNum: null,
        targetNum: null,
        workSheetArr: [],
        targetList: [],
        tableData: [],
        tableHeadName: {
          可出貨材積: 'volume',
          單位名稱: 'companyName',
          流水號: 'serialNum',
          狀態: 'state',
          產品名稱: 'productName',
          產地: 'origin',
          申請地區: 'area',
          申請編號: 'applicationNum',
          申請者姓名: 'name',
        },
      },
    };
  },
  methods: {
    // 重製
    resetData() {
      this.loading = false;
      this.targetData.currentFile = '';
      this.targetData.maxNum = null;
      this.targetData.targetNum = null;
      this.targetData.workSheetArr = [];
      this.targetData.targetList = [];
      this.targetData.tableData = [];
    },
    // 上傳前結果
    handleUpload(file) {
      this.resetData();
      this.targetData.currentFile = file.name;
      let files = { 0: file };
      this.readExcelFile(files);
      return false;
    },
    // 讀取資料
    readExcelFile(files) {
      if (files.length <= 0) {
        return;
      }
      if (!/\.(xls|xlsx)$/.test(files[0].name.toLowerCase())) {
        this.$message.error('上傳格式不正確，請上傳xls或xlsx格式');
        return;
      }
      const fileReader = new FileReader();
      fileReader.onload = (ev) => {
        try {
          const data = ev.target.result;
          const workbook = read(data, { type: 'binary' });
          // 取第一張表
          const wsname = workbook.SheetNames[0];
          // 生成JSON表格內容
          const ws = utils.sheet_to_json(workbook.Sheets[wsname]);
          this.targetData.workSheetArr = JSON.parse(JSON.stringify(ws));
          this.targetData.maxNum = ws.length;
        } catch (error) {
          return false;
        }
      };
      fileReader.readAsBinaryString(files[0]);
    },
    // 抽查數量改變
    handleChange(value) {
      this.targetData.targetNum = value;
    },
    // 抽選
    handleRandom() {
      this.loading = true;
      let currentArr = [];
      let tempWorkSheet = JSON.parse(
        JSON.stringify(this.targetData.workSheetArr)
      );
      this.targetData.targetList = [];
      this.targetData.tableData = [];
      for (let index = 0; index < this.targetData.targetNum; index++) {
        let randomNum = Math.floor(Math.random() * tempWorkSheet.length);
        currentArr.push(tempWorkSheet.splice(randomNum, 1)[0]);
      }
      this.targetData.targetList = currentArr;
      this.$message({
        message: '抽選成功!',
        type: 'success',
      });
      this.loading = false;
      currentArr.forEach((item) => {
        const newItem = {};
        Object.keys(item).forEach((key) => {
          newItem.title = key;
          newItem[this.targetData.tableHeadName[key]] = item[key];
        });
        this.targetData.tableData.push(newItem);
      });
      this.exportXlsx(this.targetData.targetList);
    },
    // 輸出xlsx
    exportXlsx(data) {
      const book = utils.book_new();
      const sheet = utils.json_to_sheet(data);
      utils.book_append_sheet(book, sheet);
      writeFile(book, `抽驗名單_${this.$moment().format('YYYY-MM-DD')}.xlsx`);
    },
  },
};
</script>

<!-- Add "scoped" attribute to limit CSS to this component only -->
<style scoped lang="scss">
.card-box {
  margin: 30px;
}
.textType {
  font-size: 14px;
  font-weight: bolder;
}
.el-upload__tip {
  margin: 20px 0;
}
.target-set > p {
  display: inline;
  margin-right: 15px;
  font-weight: bolder;
}
.upload-btn {
  background-color: #005caf;
  color: white;
}
.random-btn {
  margin-left: 1.5rem;
  background-color: #00aa90;
  color: white;
}
.title {
  font-size: 24px;
  font-weight: bolder;
}
</style>
