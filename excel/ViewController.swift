//
//  ViewController.swift
//  excel
//
//  Created by michael on 2022/2/7.
//

import Cocoa

class ViewController: NSViewController {
    @IBOutlet var textView: NSTextView!
    @IBOutlet var regTextView: NSTextView!
    @IBAction func exportExcelAction(_ sender: Any) {
        let content = textView.string
        let reg = regTextView.string // \/[a-zA-Z]{1,}\.m
        if content.isEmpty || reg.isEmpty {
            return
        }

        let muStr = getRegExpressResultString(content, regExp: reg)

        if muStr.isEmpty {
            return
        }
        saveFile(excel: muStr)
    }

    func saveFile(excel: String) {
        let openPanel = NSSavePanel()
//              openPanel.allowsMultipleSelection = false;
//              openPanel.canChooseDirectories = true;
//              openPanel.canChooseFiles = true;
        openPanel.message = "本应用需要访问该目录，请点击允许按钮"
        openPanel.prompt = "允许"
        let fm = DateFormatter()
        fm.dateFormat = "YYYY-MM-dd-HH-mm-SS"

        openPanel.nameFieldStringValue = "\(fm.string(from: Date())).xlsx"
        // openPanel.allowedFileTypes = [".xlsx"]
        openPanel.directoryURL = URL(string: NSHomeDirectory())
        openPanel.begin(completionHandler: { result in
            if result == NSApplication.ModalResponse.OK {
                //        //  文件管理器
                let fileManager = FileManager()
                // 使用 UTF16 才能显示汉字；如果显示为 ####### 是因为格子宽度不够，拉开即可
                let fileData = excel.data(using: .utf16)
//                //  文件路径
//                let path = NSHomeDirectory()
//                let filePath = URL(fileURLWithPath: path).appendingPathComponent("/Documents/export.xls").path
//                print(" 文件路径： \n\(filePath)")
                //  生成 xls 文件
                fileManager.createFile(atPath: openPanel.url!.path, contents: fileData, attributes: nil)
            }
        })
    }

    func getRegExpressResultString(_ content: String, regExp: String) -> String {
        var allResult = ""
        var regex: NSRegularExpression?
        do {
            regex = try NSRegularExpression(pattern: regExp, options: .caseInsensitive)
        } catch {
        }
        regex?.enumerateMatches(in: content, options: [], range: NSRange(location: 0, length: content.count), using: { result, _, _ in
            if let newResult = result {
                let resultRange = newResult.range(at: 0)
                var result = content.substring(with: Range(resultRange, in: content)!)
                result = "\(result)\n"
                allResult = allResult + result
            }
        })
        return allResult
    }

    override func viewDidLoad() {
        super.viewDidLoad()
        // Do any additional setup after loading the view.
    }

    override var representedObject: Any? {
        didSet {
            // Update the view, if already loaded.
        }
    }
}
