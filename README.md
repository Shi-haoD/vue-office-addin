# Vue2 + Office Add-in关于用vue项目于加载项控制excel单元格内容（Demo版）

csdn文档：https://blog.csdn.net/Shi_haoliu/article/details/153469921?spm=1001.2014.3001.5501

主要步骤：写好vue代码>生成密钥>启动代码>在window下创建一个共享文件夹>放入manifest.xml文件>打开excel的加载项>找到共享文件夹>点击加载
环境：
![在这里插入图片描述](https://i-blog.csdnimg.cn/direct/4320715476304918b90dc308e6f1b768.png)
![在这里插入图片描述](https://i-blog.csdnimg.cn/direct/de6b9ff5440b49beb507813eb82a71e9.png)
![在这里插入图片描述](https://i-blog.csdnimg.cn/direct/3e2c6360d8f74b2c973ad42a846f5736.png)
系统是win10
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
2025-10-20增加了任务窗格的图标
```javascript
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <Id>c49de0c7-98e1-420f-8e3f-cb6fb40fc047</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Excel-HelloWorld-Taskpane-JS"/>
  <Description DefaultValue="A template to get started."/>
  <IconUrl DefaultValue="https://localhost:8080/assets/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="https://localhost:8080/assets/icon-64.png"/>
  <SupportUrl DefaultValue="https://www.contoso.com/help"/>
  <AppDomains>
    <AppDomain>https://www.contoso.com</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Workbook"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://localhost:8080"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Workbook">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="CommandsGroup">
                <Label resid="CommandsGroup.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16"/>
                  <bt:Image size="32" resid="Icon.32x32"/>
                  <bt:Image size="80" resid="Icon.80x80"/>
                </Icon>
                <Control xsi:type="Button" id="TaskpaneButton">
                  <Label resid="TaskpaneButton.Label"/>
                  <Supertip>
                    <Title resid="TaskpaneButton.Label"/>
                    <Description resid="TaskpaneButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>ButtonId1</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="https://localhost:8080/assets/icon-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="https://localhost:8080/assets/icon-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="https://localhost:8080/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="GetStarted.LearnMoreUrl" DefaultValue="https://go.microsoft.com/fwlink/?LinkId=276812"/>
        <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:8080"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GetStarted.Title" DefaultValue="Get started with your sample add-in!"/>
        <bt:String id="CommandsGroup.Label" DefaultValue="Vue Test"/>
        <bt:String id="TaskpaneButton.Label" DefaultValue="Show Vue"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="GetStarted.Description" DefaultValue="Your sample add-in loaded successfully. Go to the HOME tab and click the 'Show Taskpane' button to get started."/>
        <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Click to Show a Taskpane"/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
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
exe打包工具
https://jrsoftware.org/isdl.php#stable
创建iss文件
![在这里插入图片描述](https://i-blog.csdnimg.cn/direct/c44e0678f2834651b5f9869f2c11d588.png)
在setup文件中编辑，并把manifest.xml放在当前目录
`注意这个文件中的注册表和共享文件目录位置，我用的是office365，其他版本的适配还没测试`
```javascript
; ----------------------------
; Inno Setup 完整安装器脚本（自带 GUID 生成）
; ----------------------------
[Setup]
AppName=Vue Office Addin
AppVersion=1.0.0
DefaultDirName={commonappdata}\VueOfficeAddin
OutputBaseFilename=VueOfficeAddinInstaller
Compression=lzma
SolidCompression=yes
PrivilegesRequired=admin
Uninstallable=yes
WizardStyle=modern

[Files]
Source: "manifest.xml"; DestDir: "{app}"; Flags: ignoreversion

[Code]
var
  UserChoice: Integer;

function GetComputerNameStr(): String;
begin
  Result := GetEnv('COMPUTERNAME');
  if Result = '' then
    Result := 'localhost';
end;

procedure CreateNetworkShare(ShareName, FolderPath: String);
var
  ResultCode: Integer;
begin
  if not DirExists(FolderPath) then
    CreateDir(FolderPath);

  { 使用 PowerShell 强制创建共享 }
  Exec('powershell',
       '-Command "New-SmbShare -Name ' + ShareName + ' -Path ''' + FolderPath + ''' -FullAccess Everyone"',
       '', SW_HIDE, ewWaitUntilTerminated, ResultCode);

  if ResultCode = 0 then
    MsgBox('✅ 共享文件夹已创建: \\'+GetComputerNameStr()+'\' + ShareName, mbInformation, MB_OK)
  else
    MsgBox('❌ 创建共享失败，请检查权限', mbCriticalError, MB_OK);
end;

function GenerateGUIDString(): String;
var
  i: Integer;
  HexDigits: String;
begin
  HexDigits := '0123456789ABCDEF';
  Result := '{';
  for i := 1 to 8 do Result := Result + HexDigits[Random(16)+1];
  Result := Result + '-';
  for i := 1 to 4 do Result := Result + HexDigits[Random(16)+1];
  Result := Result + '-';
  for i := 1 to 4 do Result := Result + HexDigits[Random(16)+1];
  Result := Result + '-';
  for i := 1 to 4 do Result := Result + HexDigits[Random(16)+1];
  Result := Result + '-';
  for i := 1 to 12 do Result := Result + HexDigits[Random(16)+1];
  Result := Result + '}';
end;

procedure RegisterExcelSharedFolder(ShareName: String);
var
  UNCPath, GUIDStr, RegPath: String;
  ResultCode: Integer;
begin
  UNCPath := '\\' + GetComputerNameStr() + '\' + ShareName;
  GUIDStr := GenerateGUIDString();
  RegPath := 'HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\' + GUIDStr;

  { 写 Flags }
  Exec(ExpandConstant('{sys}\reg.exe'),
       'add "' + RegPath + '" /v "Flags" /t REG_DWORD /d 1 /f',
       '', SW_HIDE, ewWaitUntilTerminated, ResultCode);

  { 写 Id }
  Exec(ExpandConstant('{sys}\reg.exe'),
       'add "' + RegPath + '" /v "Id" /t REG_SZ /d "' + GUIDStr + '" /f',
       '', SW_HIDE, ewWaitUntilTerminated, ResultCode);

  { 写 Url }
  Exec(ExpandConstant('{sys}\reg.exe'),
       'add "' + RegPath + '" /v "Url" /t REG_SZ /d "' + UNCPath + '" /f',
       '', SW_HIDE, ewWaitUntilTerminated, ResultCode);

  MsgBox('✅ Excel 共享目录已注册成功:'#13#10 +
         '路径: ' + UNCPath, mbInformation, MB_OK);
end;

procedure ForceRemoveNetworkShare(ShareName: String);
var
  ResultCode: Integer;
begin
  Exec('powershell',
       '-Command "Get-SmbShare -Name ' + ShareName + ' | Remove-SmbShare -Force"',
       '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  if ResultCode = 0 then
    MsgBox('✅ 网络共享已删除: ' + ShareName, mbInformation, MB_OK)
  else
    MsgBox('⚠ 网络共享删除失败或不存在: ' + ShareName, mbInformation, MB_OK);
end;

procedure DeleteExcelRegistry();
var
  ResultCode: Integer;
begin
  Exec(ExpandConstant('{sys}\reg.exe'),
       'delete "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs" /f',
       '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
end;

procedure DeleteSharedFolder(FolderPath: String);
begin
  if DirExists(FolderPath) then
  begin
    if DelTree(FolderPath, True, True, True) then
      MsgBox('✅ 文件夹已删除: ' + FolderPath, mbInformation, MB_OK)
    else
      MsgBox('❌ 无法删除文件夹: ' + FolderPath + '，请确保没有程序占用', mbCriticalError, MB_OK);
  end;
end;

procedure InitializeWizard();
begin
  UserChoice := MsgBox('请选择操作:'#13#10+
                       'Yes = 安装（创建共享 + 拷贝 manifest + 注册 Excel）'#13#10+
                       'No  = 卸载（删除共享 + 删除注册表 + 删除文件夹）',
                       mbConfirmation, MB_YESNO);
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  AppFolder: String;
begin
  AppFolder := ExpandConstant('{app}');

  if CurStep = ssPostInstall then
  begin
    if UserChoice = idYes then
    begin
      CreateNetworkShare('VueOfficeAddin', AppFolder);
      RegisterExcelSharedFolder('VueOfficeAddin');
      MsgBox('✅ 安装完成！', mbInformation, MB_OK);
    end
    else if UserChoice = idNo then
    begin
      ForceRemoveNetworkShare('VueOfficeAddin');
      DeleteExcelRegistry();
      DeleteSharedFolder(AppFolder);
      MsgBox('✅ 卸载完成！', mbInformation, MB_OK);
    end;
  end;
end;

```
