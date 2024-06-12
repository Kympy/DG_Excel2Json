# Install

## C# Unity

- KOR
  + Unity 폴더의 Plugins 폴더를 프로젝트 Plugins 폴더에 복사한다.
  + 에디터 실행 후, 포함된 dll 의 모든 CheckBox 를 꺼준다.
- ENG
  + Copy all 'Plugins' folder in 'Unity' folder to your project 'Plugins' folder.
  + After run editor, turn off all checkbox of dll files.

## C# exe only

- KOR
  + 빌드된 exe 파일을 실행하고 설명에 맞게 인자들을 입력한다.
  + exe 를 직접 실행하였으므로, 아마 Json 만 만들면 되는 경우 일 것이다. 인자는 2개만 입력하면 된다. 1 : Excel 폴더 경로 / 2 : Json 폴더 경로
  + 마지막으로 성공적으로 json 을 만들었다면, 마지막 인자들이 세이브파일로 저장되어 다음 번 실행 시 이어 사용할 수 있다.
 
- ENG
  + Run .exe file and enter arguments.
  + Maybe you want a only excel to json. Not C# script I think. Then you can just put only two arguments. 1 : Excel folder path / 2 : Json output folder path
  + If you success to create json file, the last arguments that you entered will save in your local space. And you can use that easily next time.
