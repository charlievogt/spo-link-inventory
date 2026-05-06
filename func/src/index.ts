// Azure Functions v4 programming model: every function file MUST be imported
// here for its app.http()/app.timer()/app.storageQueue() registration to fire.
// Adding a new functions/*.ts file without a matching import below results
// in the endpoint being silently absent at runtime.

import "./functions/linkInventoryPing.js";
import "./functions/linkInventoryScanSite.js";
import "./functions/linkInventoryScan.js";
import "./functions/linkInventoryScanStatus.js";
import "./functions/linkInventoryWhoami.js";
import "./functions/linkInventoryReplace.js";
import "./functions/linkInventoryScanWorker.js";
import "./functions/linkInventoryScanUnified.js";
import "./functions/linkInventoryScanDocs.js";
import "./functions/linkInventoryScanDocsWorker.js";
import "./functions/linkInventoryScanDocsPromote.js";
import "./functions/linkInventorySites.js";
import "./functions/linkInventoryScanDelete.js";
import "./functions/linkInventoryRetention.js";
import "./functions/linkInventoryBacklinks.js";
import "./functions/linkInventoryBacklinksBatch.js";
import "./functions/linkInventoryBacklinksExport.js";
import "./functions/linkInventoryBacklinksRebuild.js";
import "./functions/linkInventoryBacklinksColumnEnable.js";
import "./functions/duplicatesLookup.js";
import "./functions/duplicatesReport.js";
import "./functions/duplicatesExport.js";
import "./functions/duplicatesAllowlist.js";
import "./functions/duplicatesBootstrapFromVersions.js";
import "./functions/duplicatesBootstrapStatus.js";
import "./functions/duplicatesBootstrapCancel.js";
import "./functions/duplicatesBootstrapWorker.js";
import "./functions/linkInventorySchedule.js";
import "./functions/linkInventoryScheduleTimer.js";
import "./functions/orphanAssetsReport.js";
import "./functions/orphanAssetsRecycle.js";
