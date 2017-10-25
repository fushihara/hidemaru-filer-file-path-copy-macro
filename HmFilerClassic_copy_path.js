var items = [];
getSelectedItems(items);
getCurrentDirectory(items);

var varDic = new ActiveXObject("Scripting.Dictionary");
for(var i = 0 ;i<items.length;i++){
  varDic.add(varDic.count,items[i].label);
}

var menuSelectIndex = MenuArray(varDic.items());
if( menuSelectIndex == 0 ){
  // MenuArray���L�����Z��������
  EndMacro();
}
var selectObject = items[menuSelectIndex-1];
if( selectObject.copyText == ""){
  // ��؂����u���̃f�B���N�g���v�̐������Ƃ������R�s�[���Ȃ��e�L�X�g
  EndMacro();
}
SetClipboard(selectObject.copyText);

EndMacro();

function getCurrentDirectory(list){
  if(list.length!=0){
    list.push({"label":"\x01","copyText":""});
  }
  list.push({"label":"���̃f�B���N�g��","copyText":""});
  var nowDirectory = GetDirectory();
  var items = getPathFormatList(GetDirectory());
  for(var i =0;i<items.length;i++){
    list.push({"label":items[i],"copyText":items[i]});
  }
}
function getSelectedItems(list){
  if( GetSelectedCount() == 0 ){
    return;
  }
  var items = [];
  var index = -1;
  var fileCountFile   = 0;
  var fileCountFolder = 0;
  while(true){
    var index=GetNextItem(index,0x02);
    if(index==-1){break;}
    var path     = GetItemPath(index);
    var isFolder = IsFolder(index);
    items.push(getPathFormatList(path));
    if( isFolder ){
      fileCountFolder += 1;
    }else{
      fileCountFile   += 1;
    }
  }
  list.push({"label":"�I�𒆂�"+fileCountFile+"�̃t�@�C���A"+fileCountFolder+"�̃t�H���_(���v"+(fileCountFile+fileCountFolder)+"��)","copyText":""});
  for(var i = 0;i<items[0].length;i++){
    var label = items[0][i];
    var copyTextList = [];
    for(var j = 0;j<items.length;j++){
      copyTextList.push(items[j][i]);
    }
    list.push({"label":label,"copyText":copyTextList.join("\n")});
  }
}
// �n���ꂽ�t�@�C���p�X���A�_�u���N�I�[�g�ň͂�Ȃ��A�͂��Abash on windows�̃t�H�[�}�b�g�Ń_�u���N�I�[�g�ň͂�Ȃ��A�͂��`���ɕϊ����Ĕz��ŕԂ�
function getPathFormatList(path){
  var result = [];
  var pathEscaped = path.replace(/\\/g,"\\\\");
  result.push(path);
  result.push("\""+path+"\"");
  result.push(pathEscaped);
  result.push("\""+pathEscaped+"\"");
  result.push("/mnt/"   + path.replace(/(\w):\\/g,function(a,b){return b.toLowerCase()+"/";}).replace(/\\/g,"/"));
  result.push("\"/mnt/" + path.replace(/(\w):\\/g,function(a,b){return b.toLowerCase()+"/";}).replace(/\\/g,"/")+"\"");
  return result;
}
