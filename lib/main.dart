import 'dart:convert';
import 'dart:io';
import 'dart:isolate';

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
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Flutter Demo',
      theme: ThemeData(
        colorScheme: ColorScheme.fromSeed(seedColor: Colors.deepPurple),
        useMaterial3: true,
      ),
      home: const MyHomePage(title: 'PPTX UNLOCK'),
    );
  }
}

class MyHomePage extends StatefulWidget {
  const MyHomePage({super.key, required this.title});

  final String title;

  @override
  State<MyHomePage> createState() => _MyHomePageState();
}

class _MyHomePageState extends State<MyHomePage> {
  void _isolateDecodeUtf(SendPort port) {
    final ReceivePort res = ReceivePort();
    port.send(res.sendPort);
    res.listen((message) {
      final SendPort send = message[0] as SendPort;
      final List<int> data = message[1] as List<int>;
      send.send(utf8.decode(data));
    });
  }

  void _pickXmlFile() async {
    FilePickerResult? result = await FilePicker.platform.pickFiles();

    if (result != null) {
      final path = result.files.single.path!;
      // 형식이 맞을 때
      if (path.endsWith('presentation.xml')) {
        File file = File(result.files.single.path!);
        // inspect(file);
        final document = XmlDocument.parse(file.readAsStringSync());
        final el = document.findAllElements('p:modifyVerifier');
        // inspect(document);

        for (var element in el) {
          element.remove();
        }
        // print(document.toXmlString());
        file.writeAsStringSync(document.toXmlString());

        // 완료
        _openDialog('완료', '닫기');
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
        var temp = path.split('.');
        temp.last = 'zip';
        String newPath = temp.join('.');
        print('newPath: $newPath');
        file = file.renameSync(newPath);
        final bytes = await file.readAsBytes();
        final archive = ZipDecoder().decodeBytes(bytes);
        final presentation = archive.files.firstWhereOrNull(
            (element) => element.name == 'ppt/presentation.xml');
        if (presentation != null) {
          // await Isolate.spawn((message) {

          // }, message)
          final content = utf8.decode(presentation.content as List<int>);
          final xml = XmlDocument.parse(content);
          final pModify = xml.findAllElements('p:modifyVerifier');

          for (var el in pModify) {
            el.remove();
          }

          // final editedArchive = Archive();
          // for (var file in archive.files) {
          //   if (file != presentation) {
          //     editedArchive.addFile(file);
          //   }
          // }

          final editedContent = xml.toXmlString();
          final editedBytes = utf8.encode(editedContent);

          final editedPresentation = ArchiveFile(
              'ppt/presentation.xml', editedBytes.length, editedBytes);

          archive.addFile(editedPresentation);
          final editedArchiveBytes = ZipEncoder().encode(archive);

          if (editedArchiveBytes != null) {
            final editedFile = File(path);
            await editedFile.writeAsBytes(editedArchiveBytes);
          }
        }
      }
      // 형식이 안 맞을 때
      else {}
    }
    // 취소
    else {}
  }

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
        actions: [
          ElevatedButton(
              onPressed: () => Navigator.of(context).pop(),
              child: Text(buttonText))
        ],
      ),
    );
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
                  'presentaiton.xml 파일 선택',
                ))
          ],
        ),
      ),
    );
  }
}
