import 'dart:convert';
import 'dart:io';

import 'package:archive/archive_io.dart';
import 'package:desktop_drop/desktop_drop.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';
import 'package:archive/archive.dart';
import 'package:xml/xml.dart';

void main() {
  runApp(const App());
}

class App extends StatelessWidget {
  const App({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'UNLOCK-PPTX',
      theme: ThemeData(
        colorScheme: ColorScheme.fromSeed(seedColor: Colors.deepPurple),
        useMaterial3: true,
      ),
      home: const Home(title: 'UNLOCK-PPTX'),
    );
  }
}

class Home extends StatefulWidget {
  const Home({super.key, required this.title});

  final String title;

  @override
  State<Home> createState() => _HomeState();
}

class _HomeState extends State<Home> {
  // 드래그 중인지
  bool isDragging = false;

  Future<void> _openDialog(String title, String buttonText) async {
    await showDialog(
      context: context,
      builder: (context) => AlertDialog(
        alignment: Alignment.center,
        title: Text(
          title,
          textAlign: TextAlign.center,
        ),
        titlePadding: const EdgeInsets.all(30),
        actionsAlignment: MainAxisAlignment.center,
        actions: [ElevatedButton(onPressed: () => Navigator.of(context).pop(), child: Text(buttonText))],
      ),
    );
  }

  // xml 수정 가능하게 편집
  // bool _convertXmlToEdit(String path) {
  //   try {
  //     File file = File(path);
  //     // inspect(file);
  //     final document = XmlDocument.parse(file.readAsStringSync());
  //     final el = document.findAllElements('p:modifyVerifier');
  //     // inspect(document);

  //     for (var element in el) {
  //       element.remove();
  //     }
  //     // print(document.toXmlString());
  //     file.writeAsStringSync(document.toXmlString());

  //     return true;
  //   } catch (e) {
  //     return false;
  //   }
  // }

  // xml 파일 선택
  // void _pickXmlFile() async {
  //   FilePickerResult? result = await FilePicker.platform.pickFiles();

  //   if (result != null) {
  //     final path = result.files.single.path!;
  //     // 형식이 맞을 때
  //     if (path.endsWith('presentation.xml')) {
  //       final result = _convertXmlToEdit(path);

  //       if (result) {
  //         // 완료
  //         _openDialog('완료', '닫기');
  //       } else {
  //         _openDialog('저장에 실패했습니다', '닫기');
  //       }
  //     }
  //     // 형식이 안 맞을 때
  //     else {
  //       _openDialog('잘못된 파일 입니다', '닫기');
  //     }
  //   }
  //   // 취소
  //   else {}
  // }

  // path 받아서 처리
  void _convertProcessByPath(String path) async {
    // 형식이 맞을 때
    if (path.endsWith('pptx')) {
      final result = await _convertPptxToEdit(path);

      if (result) {
        // 완료
        _openDialog('완료', '닫기');
      }
    }
    // 형식이 안 맞을 때
    else {
      _openDialog('잘못된 파일 입니다', '닫기');
    }
  }

  // TODO buffer 사용하는 방식으로 변경, zip 저장하지 않고 하는 방법, 로딩 인디게이터
  // pptx 파일 수정
  Future<bool> _convertPptxToEdit(String path) async {
    try {
      // 기존 파일
      File file = File(path);
      final bytes = await file.readAsBytes();
      // -1.zip 파일 생성
      final temp = path.split('.');
      temp[temp.length - 2] += '-1';
      temp.last = 'zip';
      String newPath = temp.join('.');
      File newFile = File(newPath);
      newFile.writeAsBytesSync(bytes);
      // 작업 시작
      final newBytes = newFile.readAsBytesSync();
      final archive = ZipDecoder().decodeBytes(newBytes, verify: true);
      final editedArchive = Archive(); // 새로 저장할 용도
      for (var file in archive) {
        // presentation.xml 찾기
        if (file.name == 'ppt/presentation.xml') {
          String result = utf8.decode(file.content);
          final document = XmlDocument.parse(result);
          final el = document.findAllElements('p:modifyVerifier');

          for (var element in el) {
            element.remove();
          }

          // 변환 후 추가
          final xml = document.toXmlString();
          final xmlBytes = utf8.encode(xml);
          final newXml = ArchiveFile('ppt/presentation.xml', xmlBytes.length, xmlBytes);
          editedArchive.addFile(newXml);
        }
        // 나머지 파일은 바로 추가
        else {
          editedArchive.addFile(file);
        }
      }

      // -1.pptx로 변환
      final editedBytes = ZipEncoder().encode(editedArchive);
      if (editedBytes != null) {
        newFile.writeAsBytesSync(editedBytes);
      }
      temp.last = 'pptx';
      newPath = temp.join('.');
      newFile.renameSync(newPath);

      // 저장 성공
      return true;
    } catch (error) {
      await _openDialog('변환 도중 실패: $error', '확인');
      return false;
    }
  }

  // 파일 선택
  void _pickPptxFile() async {
    FilePickerResult? pickedFiles = await FilePicker.platform.pickFiles();

    if (pickedFiles != null) {
      final path = pickedFiles.files.single.path!;
      // 처리
      _convertProcessByPath(path);
    }
  }

  // pptx 드래그 앤 드랍
  void _onDragDropPptxFile(DropDoneDetails details) async {
    if (details.files.isNotEmpty) {
      final path = details.files.first.path;
      // 처리
      _convertProcessByPath(path);
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      restorationId: 'unlock-pptx',
      appBar: AppBar(
        backgroundColor: Theme.of(context).colorScheme.inversePrimary,
        title: Text(widget.title),
      ),
      body: Center(
        child: Column(
          mainAxisAlignment: MainAxisAlignment.center,
          children: <Widget>[
            // Padding(
            //   padding: const EdgeInsets.all(30),
            //   child: Text(
            //     'pptx만 넣으면 되게 지원 예정',
            //     style: Theme.of(context).textTheme.headlineMedium,
            //   ),
            // ),
            // ElevatedButton(
            //     onPressed: _pickXmlFile,
            //     child: const Text(
            //       'presentation.xml 파일 선택',
            //     )),
            ElevatedButton(
                onPressed: _pickPptxFile,
                child: const Text(
                  'pptx 파일 선택',
                )),
            DropTarget(
                onDragDone: _onDragDropPptxFile,
                onDragEntered: (details) {
                  setState(() {
                    isDragging = true;
                  });
                },
                onDragExited: (details) {
                  setState(() {
                    isDragging = false;
                  });
                },
                child: Padding(
                  padding: const EdgeInsets.all(40),
                  child: Container(
                    width: double.infinity,
                    padding: const EdgeInsets.symmetric(vertical: 30),
                    alignment: Alignment.center,
                    decoration: BoxDecoration(
                        border: Border.all(color: isDragging ? Colors.deepPurple : Colors.grey, width: 3),
                        borderRadius: BorderRadius.circular(20)),
                    child: RichText(
                      text: TextSpan(children: [
                        TextSpan(
                          text: 'Drag&Drop ',
                          style: TextStyle(
                            fontSize: 18,
                            color: isDragging ? Colors.deepPurple : Colors.grey,
                          ),
                        ),
                        TextSpan(
                          text: '\'.pptx\'',
                          style: TextStyle(
                            fontSize: 18,
                            fontWeight: FontWeight.bold,
                            color: isDragging ? Colors.deepPurple : Colors.grey,
                          ),
                        )
                      ]),
                    ),
                    // child: Text(
                    //   'Drag&Drop \'presentation.xml\'',
                    //   style: TextStyle(
                    //     fontSize: 18,
                    //     color: isDragging ? Colors.deepPurple : Colors.grey,
                    //   ),
                    // ),
                  ),
                )),
          ],
        ),
      ),
    );
  }
}
