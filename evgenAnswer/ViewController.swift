//
//  ViewController.swift
//  evgenAnswer
//
//  Created by Ilya Aleshin on 26/03/2019.
//  Copyright © 2019 iAleshin. All rights reserved.
//

import UIKit
import SwiftyDropbox
import CoreXLSX

class ViewController: UIViewController {
    
    var client: DropboxClient?
    
    @IBOutlet weak var downloadBtn: UIButton!
    override func viewDidLoad() {
        super.viewDidLoad()
        let token: String = UserDefaults.standard.string(forKey: "db_token") ?? ""
        if token != "" {
            client = DropboxClient(accessToken: token)
            downloadBtn.isHidden = false
        }
        // Do any additional setup after loading the view, typically from a nib.
    }
    
    func dbReturnSuccess(url: URL) {
        client = DropboxClientsManager.authorizedClient
        UserDefaults.standard.setValue(client?.auth.client.accessToken, forKey: "db_token")
        downloadBtn.isHidden = false
    }

    @IBAction func downloadAction(_ sender: Any) {
        client!.files.download(path: "/evgenAnswer/example.xlsx")
            .response { response, error in
                if let response = response {
                    print("--- downloaded!")
                    let responseMetadata = response.0
                    print(responseMetadata)
                    let fileContents = response.1
                    print(fileContents)
                    let path = self.saveFile(d: fileContents)
                    self.readExcel(path: path)
                } else if let error = error {
                    print(error)
                }
            }
            .progress { progressData in
                print(progressData)
        }
    }
    
    func readExcel(path: String) {
        let file = XLSXFile(filepath: path)
        let sharedStrings = try! file!.parseSharedStrings()
        for path in try! file?.parseWorksheetPaths() ?? [] {
            let ws = try! file?.parseWorksheet(at: path)
            let columnCStrings = ws!.cells(atColumns: [ColumnReference("C")!])
                .compactMap { $0.value }
                .compactMap { Int($0) }
                .compactMap { sharedStrings.items[$0].text }
            let rows = ws!.cells(atColumns: [ColumnReference("C")!])
                .compactMap { $0.value }
                .compactMap { Int($0) }
            for s in rows {
                print("row=", s)
            }
            for s in columnCStrings {
                print("s=", s)
            }
        }
    }
    
    func saveFile(d: Data) -> String {
        let fileName = "Test"
        let DocumentDirURL = try! FileManager.default.url(for: .documentDirectory, in: .userDomainMask, appropriateFor: nil, create: true)
        let fileURL = DocumentDirURL.appendingPathComponent(fileName).appendingPathExtension("xlsx")
        try! d.write(to: fileURL, options: .noFileProtection)
        return fileURL.path
    }
    
    
    @IBAction func startAction(_ sende: Any) {
        DropboxClientsManager.authorizeFromController(UIApplication.shared,
                                                      controller: self,
                                                      openURL: { (url: URL) -> Void in
//                                                        UIApplication.shared.openURL(url)
                                                        UIApplication.shared.open(url, options: [:], completionHandler: { (r) in
                                                            print("flag")
                                                        })
        })

    }
    
}

