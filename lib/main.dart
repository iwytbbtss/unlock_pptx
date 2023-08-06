import 'dart:convert';
import 'dart:io';
import 'dart:isolate';

import 'package:archive/archive_io.dart';
import 'package:desktop_drop/desktop_drop.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';
import 'package:archive/archive.dart';
import 'package:xml/xml.dart';

extension IterableX<T> on Iterable<T> {
  T? firstWhereOrNull(bool Function(T element) test) {
    if (isEmpty) {
      return null;
    }

    for (var el in this) {
      if (test(el)) return el;
    }

    return null;
  }
}

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
      home: const Home(title: 'UNLOCK PPTX'),
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

  void _openDialog(String title, String buttonText) {
    showDialog(
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
  bool _convertXmlToEdit(String path) {
    try {
      File file = File(path);
      // inspect(file);
      final document = XmlDocument.parse(file.readAsStringSync());
      final el = document.findAllElements('p:modifyVerifier');
      // inspect(document);

      for (var element in el) {
        element.remove();
      }
      // print(document.toXmlString());
      file.writeAsStringSync(document.toXmlString());

      return true;
    } catch (e) {
      return false;
    }
  }

  // xml 파일 선택
  void _pickXmlFile() async {
    FilePickerResult? result = await FilePicker.platform.pickFiles();

    if (result != null) {
      final path = result.files.single.path!;
      // 형식이 맞을 때
      if (path.endsWith('presentation.xml')) {
        final result = _convertXmlToEdit(path);

        if (result) {
          // 완료
          _openDialog('완료', '닫기');
        } else {
          _openDialog('저장에 실패했습니다', '닫기');
        }
      }
      // 형식이 안 맞을 때
      else {
        _openDialog('잘못된 파일 입니다', '닫기');
      }
    }
    // 취소
    else {}
  }

  void _pickPptxFile() async {
    FilePickerResult? result = await FilePicker.platform.pickFiles();

    if (result != null) {
      final path = result.files.single.path!;
      // 형식이 맞을 때
      if (path.endsWith('pptx')) {
        File file = File(result.files.single.path!);
        final temp = path.split('.');
        temp[temp.length - 2] += '-1';
        temp.last = 'zip';
        String newPath = temp.join('.');
        print('newPath: $newPath');
        final bytes = await file.readAsBytes();
        // -1.zip 생성
        File newFile = File(newPath);
        newFile.writeAsBytesSync(bytes);
        /* 여기까지 .pptx를 .zip으로 변경 */

        final inputStream = InputFileStream(newPath);
        final archive = ZipDecoder().decodeBuffer(inputStream);
        final editedArchive = Archive();
        for (var file in archive) {
          // print(file.name);
          if (file.name == 'ppt/presentation.xml') {
            String result = utf8.decode(file.content);
            print(result);
            final document = XmlDocument.parse(result);
            final el = document.findAllElements('p:modifyVerifier');

            for (var element in el) {
              element.remove();
            }

            final xml = document.toXmlString();
            final xmlBytes = utf8.encode(xml);
            final newXml = ArchiveFile('ppt/presentation.xml', xmlBytes.length, xmlBytes);
            editedArchive.addFile(newXml);
          } else {
            editedArchive.addFile(file);
          }
        }

        await inputStream.close();

        final editedBytes = ZipEncoder().encode(editedArchive);
        if (editedBytes != null) {
          newFile.writeAsBytesSync(editedBytes);
        }
        temp.last = 'pptx';
        newPath = temp.join('.');
        newFile.renameSync(newPath);
        // final presentation = archive.files.firstWhereOrNull((element) => element.name == 'ppt/presentation.xml');
        // if (presentation != null) {
        // await Isolate.spawn((message) {

        // }, message)
        // final content = utf8.decode(presentation.content as List<int>);
        // final xml = XmlDocument.parse(content);
        // final pModify = xml.findAllElements('p:modifyVerifier');

        // for (var el in pModify) {
        //   el.remove();
        // }

        // final editedArchive = Archive();
        // for (var file in archive.files) {
        //   if (file != presentation) {
        //     editedArchive.addFile(file);
        //   }
        // }

        // final editedContent = xml.toXmlString();
        // final editedBytes = utf8.encode(editedContent);

        // final editedPresentation = ArchiveFile('ppt/presentation.xml', editedBytes.length, editedBytes);

        // archive.addFile(editedPresentation);
        // final editedArchiveBytes = ZipEncoder().encode(archive);

        // if (editedArchiveBytes != null) {
        //   final editedFile = File(path);
        //   await editedFile.writeAsBytes(editedArchiveBytes);
        // }
        // }
      }
      // 형식이 안 맞을 때
      else {}
    }
    // 취소
    else {}
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
            Padding(
              padding: const EdgeInsets.all(30),
              child: Text(
                'pptx만 넣으면 되게 지원 예정',
                style: Theme.of(context).textTheme.headlineMedium,
              ),
            ),
            ElevatedButton(
                onPressed: _pickXmlFile,
                child: const Text(
                  'presentation.xml 파일 선택',
                )),
            ElevatedButton(
                onPressed: _pickPptxFile,
                child: const Text(
                  'pptx 파일 선택',
                )),
            DropTarget(
                onDragDone: (details) {
                  if (details.files.isNotEmpty) {
                    final path = details.files.first.path;
                    // 형식이 맞을 때
                    if (path.endsWith('presentation.xml')) {
                      final result = _convertXmlToEdit(path);
                      if (result) {
                        // 완료
                        _openDialog('완료', '닫기');
                      } else {
                        _openDialog('저장에 실패했습니다', '닫기');
                      }
                    }
                    // 형식이 안 맞을 때
                    else {
                      _openDialog('잘못된 파일 입니다', '닫기');
                    }
                  }
                },
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
                          text: '\'presentation.xml\'',
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
