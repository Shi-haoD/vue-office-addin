# vue-Office-Addin-DEMO
主要步骤：写好vue代码>生成密钥>启动代码>在window下创建一个共享文件夹>放入manifest.xml文件>打开excel的加载项>找到共享文件夹>点击加载
# 开发使用步骤
## 1.新建vue2结构
![在这里插入图片描述](https://i-blog.csdnimg.cn/direct/236e668e60cc4940b383eea0f0278e61.png)
### public
#### index.html
`注：这里用cdn的包，不用npm下载的包，那个我下载没成功`

```javascript
<!DOCTYPE html>
<html lang="zh">

<head>
  <meta charset="UTF-8" />
  <title>Vue2 Office Add-in</title>
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
</head>

<body>
  <div id="app"></div>
</body>

</html>
```

### src
#### components
##### HelloExcel.vue
`主页代码代码`
```javascript
<template>
	<div class="hello-excel">
		<h2>Vue2 + Office Add-in</h2>
		<button @click="writeData">写入 Excel</button>
		<button @click="readData">读取 Excel</button>
		<p v-if="cellValue">A1 内容：{{ cellValue }}</p>
	</div>
</template>

<script>
import { writeCell, readCell } from '../office';

export default {
	data() {
		return { cellValue: '' };
	},
	methods: {
		async writeData() {
			await writeCell('Hello from Vue!');
		},
		async readData() {
			this.cellValue = await readCell();
		},
	},
};
</script>

<style scoped>
.hello-excel {
	padding: 20px;
	font-family: Arial;
}
button {
	margin-right: 10px;
	padding: 6px 12px;
	cursor: pointer;
}
</style>

```

#### App.vue
`vue入口`
```javascript
<template>
	<div id="app">
		<HelloExcel />
	</div>
</template>

<script>
import HelloExcel from './components/HelloExcel.vue';

export default {
	components: { HelloExcel },
};
</script>

```

#### main.js

```javascript
import Vue from 'vue';
import App from './App.vue';

Vue.config.productionTip = false;

// Office.js 加载完成后再挂载 Vue
Office.onReady(() => {
	new Vue({
		render: (h) => h(App),
	}).$mount('#app');
});

```
`入口`
#### office.js
`这个是调用excel的单元格代码`
```javascript
/* src/office.js */
export async function writeCell(value = 'Hello Vue!') {
	await Excel.run(async (context) => {
		const sheet = context.workbook.worksheets.getActiveWorksheet();
		sheet.getRange('A1').values = [[value]];
		await context.sync();
	});
}

export async function readCell() {
	return await Excel.run(async (context) => {
		const sheet = context.workbook.worksheets.getActiveWorksheet();
		const range = sheet.getRange('A1');
		range.load('values');
		await context.sync();
		return range.values[0][0];
	});
}

```
### localhost+1-key.pem
这个是生成的密钥，需要https才能访问到页面
### localhost+1.pem
这个是生成的密钥，需要https才能访问到页面
### manifest.xml
`这个是要放在共享文件夹中的主要代码，DefaultValue="https://localhost:8080"这里是页面的地址，千万注意要是https的，否则excel那边打不开`

```javascript
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:type="TaskPaneApp">

  <Id>12345678-aaaa-bbbb-cccc-123456789abc</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Vue Office Addin</ProviderName>
  <DefaultLocale>zh-CN</DefaultLocale>
  <DisplayName DefaultValue="Vue Office Addin"/>
  <Description DefaultValue="A Vue2 Excel Add-in demo"/>

  <Hosts>
    <Host Name="Workbook"/>
  </Hosts>

  <DefaultSettings>
    <SourceLocation DefaultValue="https://localhost:8080"/>
    
  </DefaultSettings>

  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>

```
### mkcert.exe这个是下载的生成密钥的工具
### package-lock.json这个不用管
### package.json这个也没啥

```javascript
{
	"name": "vue-office-addin",
	"version": "1.0.0",
	"private": true,
	"scripts": {
		"serve": "vue-cli-service serve",
		"build": "vue-cli-service build"
	},
	"dependencies": {
		"core-js": "^3.8.3",
		"vue": "^2.7.14"
	},
	"devDependencies": {
		"@vue/cli-service": "^5.0.0",
		"vue-template-compiler": "^2.7.14"
	}
}

```

### vue.config.js
`key: fs.readFileSync(path.join(__dirname, 'localhost+1-key.pem')),
cert: fs.readFileSync(path.join(__dirname, 'localhost+1.pem')), 
			注意这两个 这两个是生成的密钥地址`

```javascript
const fs = require('fs');
const path = require('path');
module.exports = {
	devServer: {
		https: {
			key: fs.readFileSync(path.join(__dirname, 'localhost+1-key.pem')),
			cert: fs.readFileSync(path.join(__dirname, 'localhost+1.pem')),
		},
		port: 8080,
	},
};
```
## 2.生成密钥
https://github.com/FiloSottile/mkcert/releases
1. 去这个地址下载mkcert的windwos64位版本
2. 下载后把它改名为 mkcert.exe
3. 下载之后放到项目根目录下
4. 在右键文件夹中打开cmd控制台输入命令mkcert -install
5. 应该会出现
Created a new local CA
The local CA is now installed in the system trust store!
6. 之后输入命令 mkcert localhost 127.0.0.1
会生成![在这里插入图片描述](https://i-blog.csdnimg.cn/direct/10969b3c1e914c9dbbc6e71b798b19f6.png)
这两个文件，名字可能和我的不一样，记得在配置文件中改

## 3.npm下载
`npm install`
`npm run serve`
项目启动成功之后访问  https://localhost:8080/index.html
这样vue端的就完成了
## 4.excel部分
1. 新建一个文件夹，名字随意，位置随意
![在这里插入图片描述](https://i-blog.csdnimg.cn/direct/416985e222c04628ad2ade88d49ca7e6.png)
2. 创建一个共享文件夹
按着图片的流程右键文件夹属性>共享>共享，之后复制地址
![在这里插入图片描述](https://i-blog.csdnimg.cn/direct/7bec0ff19e224949bf29cb43bd8f7fb9.png)
3. 打开excel
找到excel的选项设置页，打开信任中心
![在这里插入图片描述](https://i-blog.csdnimg.cn/direct/fb2490fd9b924d97a6b805813dc8c5d0.png)
输入刚才的共享文件地址，勾选显示在菜单中
![在这里插入图片描述](https://i-blog.csdnimg.cn/direct/f0757795bee9492c90ec962ff19a3bde.png)
4. 这样再重启excel之后再点开加载项就有了
5.![在这里插入图片描述](https://i-blog.csdnimg.cn/direct/2802c295322746e39bfffa954ae3ba5a.png)
## 5.制作安装包，给客户直接一键安装好
制作中，预计使用.exe安装，然后拿注册表写到系统里
