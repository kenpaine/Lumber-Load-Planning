import UIKit
import Capacitor
import WebKit

@UIApplicationMain
class AppDelegate: UIResponder, UIApplicationDelegate {

    var window: UIWindow?

    func application(_ application: UIApplication, didFinishLaunchingWithOptions launchOptions: [UIApplication.LaunchOptionsKey: Any]?) -> Bool {
        return true
    }

    func applicationWillResignActive(_ application: UIApplication) {
        // Sent when the application is about to move from active to inactive state. This can occur for certain types of temporary interruptions (such as an incoming phone call or SMS message) or when the user quits the application and it begins the transition to the background state.
        // Use this method to pause ongoing tasks, disable timers, and invalidate graphics rendering callbacks. Games should use this method to pause the game.
    }

    func applicationDidEnterBackground(_ application: UIApplication) {
        // Use this method to release shared resources, save user data, invalidate timers, and store enough application state information to restore your application to its current state in case it is terminated later.
        // If your application supports background execution, this method is called instead of applicationWillTerminate: when the user quits.
    }

    func applicationWillEnterForeground(_ application: UIApplication) {
        // Called as part of the transition from the background to the active state; here you can undo many of the changes made on entering the background.
    }

    func applicationDidBecomeActive(_ application: UIApplication) {
        // Restart any tasks that were paused (or not yet started) while the application was inactive. If the application was previously in the background, optionally refresh the user interface.
    }

    func applicationWillTerminate(_ application: UIApplication) {
        // Called when the application is about to terminate. Save data if appropriate. See also applicationDidEnterBackground:.
    }

    func application(_ app: UIApplication, open url: URL, options: [UIApplication.OpenURLOptionsKey: Any] = [:]) -> Bool {
        // Called when the app was launched with a url. Feel free to add additional processing here,
        // but if you want the App API to support tracking app url opens, make sure to keep this call
        return ApplicationDelegateProxy.shared.application(app, open: url, options: options)
    }

    func application(_ application: UIApplication, continue userActivity: NSUserActivity, restorationHandler: @escaping ([UIUserActivityRestoring]?) -> Void) -> Bool {
        // Called when the app was launched with an activity, including Universal Links.
        // Feel free to add additional processing here, but if you want the App API to support
        // tracking app url opens, make sure to keep this call
        return ApplicationDelegateProxy.shared.application(application, continue: userActivity, restorationHandler: restorationHandler)
    }

}

// Capacitor bridge subclass — explicitly registers the app-local print plugin
// (the reliable path; app-target plugins aren't always auto-discovered). Set as the
// view controller's custom class in Main.storyboard.
class MainViewController: CAPBridgeViewController {
    override func capacitorDidLoad() {
        bridge?.registerPluginInstance(LumberPrintPlugin())
    }
}

// Native AirPrint bridge. WKWebView ignores window.print(), so the Print button
// calls this; it hands the live web view to UIPrintInteractionController, which
// renders with the page's @media print CSS — same one-page manifest as the web.
@objc(LumberPrintPlugin)
public class LumberPrintPlugin: CAPPlugin, CAPBridgedPlugin {
    public let identifier = "LumberPrintPlugin"
    public let jsName = "LumberPrint"
    public let pluginMethods: [CAPPluginMethod] = [
        CAPPluginMethod(name: "print", returnType: CAPPluginReturnPromise),
        CAPPluginMethod(name: "sharePdf", returnType: CAPPluginReturnPromise)
    ]

    @objc func print(_ call: CAPPluginCall) {
        DispatchQueue.main.async {
            guard let webView = self.bridge?.webView else {
                call.reject("No web view available to print")
                return
            }
            let controller = UIPrintInteractionController.shared
            let info = UIPrintInfo(dictionary: nil)
            info.outputType = .general
            info.jobName = call.getString("jobName") ?? "Load Manifest"
            info.orientation = .landscape
            controller.printInfo = info
            controller.printFormatter = webView.viewPrintFormatter()
            controller.present(animated: true) { (_, completed, error) in
                if let error = error {
                    call.reject(error.localizedDescription)
                } else {
                    call.resolve(["completed": completed])
                }
            }
        }
    }

    // Render the web view to a PDF (same @media print layout) and open the iOS share sheet,
    // so the user can email it (Mail attaches the PDF), Save to Files, AirDrop, etc.
    @objc func sharePdf(_ call: CAPPluginCall) {
        let subject = call.getString("subject") ?? "Load Manifest"
        let filename = call.getString("filename") ?? "LoadManifest.pdf"
        DispatchQueue.main.async {
            guard let webView = self.bridge?.webView else {
                call.reject("No web view available to export")
                return
            }
            let renderer = UIPrintPageRenderer()
            renderer.addPrintFormatter(webView.viewPrintFormatter(), startingAtPageAt: 0)
            // US Letter landscape (72 pt/in): 792 × 612, 0.25in margins
            let paper = CGRect(x: 0, y: 0, width: 792, height: 612)
            renderer.setValue(paper, forKey: "paperRect")
            renderer.setValue(paper.insetBy(dx: 18, dy: 18), forKey: "printableRect")
            let pdfData = NSMutableData()
            UIGraphicsBeginPDFContextToData(pdfData, paper, nil)
            let pages = max(renderer.numberOfPages, 1)
            for i in 0..<pages {
                UIGraphicsBeginPDFPage()
                renderer.drawPage(at: i, in: UIGraphicsGetPDFContextBounds())
            }
            UIGraphicsEndPDFContext()
            let url = FileManager.default.temporaryDirectory.appendingPathComponent(filename)
            do { try pdfData.write(to: url) } catch {
                call.reject("Could not write the PDF: \(error.localizedDescription)")
                return
            }
            let activity = UIActivityViewController(activityItems: [url], applicationActivities: nil)
            activity.setValue(subject, forKey: "subject")   // prefills the Mail subject
            if let pop = activity.popoverPresentationController {
                pop.sourceView = webView
                pop.sourceRect = CGRect(x: webView.bounds.midX, y: webView.bounds.midY, width: 0, height: 0)
                pop.permittedArrowDirections = []
            }
            self.bridge?.viewController?.present(activity, animated: true) {
                call.resolve(["shared": true])
            }
        }
    }
}
