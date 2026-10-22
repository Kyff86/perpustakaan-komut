// ==========================================================
// api.gs - gerbang API untuk versi GitHub Pages.
// Halaman di GitHub memanggil web app ini lewat fetch (POST, JSON).
// Hanya fungsi pada daftar putih API_MAP yang dapat dipanggil dari luar.
// ==========================================================
const API_MAP = {
  apiAdminUpdateUser: apiAdminUpdateUser,
  apiAuthenticateUser: apiAuthenticateUser,
  apiChangePassword: apiChangePassword,
  apiDeleteFormQuestion: apiDeleteFormQuestion,
  apiDeleteJenisInsiden: apiDeleteJenisInsiden,
  apiDeleteLampiran: apiDeleteLampiran,
  apiDeleteUser: apiDeleteUser,
  apiDownloadLampiran: apiDownloadLampiran,
  apiGetAudit: apiGetAudit,
  apiGetDashboardData: apiGetDashboardData,
  apiGetFormQuestions: apiGetFormQuestions,
  apiGetJenisInsiden: apiGetJenisInsiden,
  apiGetLampiran: apiGetLampiran,
  apiGetMyProfile: apiGetMyProfile,
  apiGetMySignature: apiGetMySignature,
  apiGetRCAByReportId: apiGetRCAByReportId,
  apiGetReviewDireksi: apiGetReviewDireksi,
  apiGetSignatures: apiGetSignatures,
  apiGetUsers: apiGetUsers,
  apiKegDelete: apiKegDelete,
  apiKegList: apiKegList,
  apiKegSave: apiKegSave,
  apiKomiteHome: apiKomiteHome,
  apiLibDeleteFile: apiLibDeleteFile,
  apiLibDeleteFolder: apiLibDeleteFolder,
  apiLibList: apiLibList,
  apiLibSaveFolder: apiLibSaveFolder,
  apiLibUpload: apiLibUpload,
  apiLogout: apiLogout,
  apiMutuDelete: apiMutuDelete,
  apiMutuList: apiMutuList,
  apiMutuSave: apiMutuSave,
  apiRegisterUser: apiRegisterUser,
  apiRemoveSignature: apiRemoveSignature,
  apiReorderFormQuestions: apiReorderFormQuestions,
  apiRequestPasswordReset: apiRequestPasswordReset,
  apiResetPassword: apiResetPassword,
  apiRestoreSession: apiRestoreSession,
  apiSaveFormQuestion: apiSaveFormQuestion,
  apiSaveJenisInsiden: apiSaveJenisInsiden,
  apiSaveMySignature: apiSaveMySignature,
  apiSignDocument: apiSignDocument,
  apiSopAction: apiSopAction,
  apiSopCreate: apiSopCreate,
  apiSopList: apiSopList,
  apiSopLogs: apiSopLogs,
  apiSubmitGradingKaru: apiSubmitGradingKaru,
  apiSubmitLaporanIKP: apiSubmitLaporanIKP,
  apiSubmitRCA: apiSubmitRCA,
  apiSubmitRegradingKomite: apiSubmitRegradingKomite,
  apiSubmitReviewDireksi: apiSubmitReviewDireksi,
  apiUpdateMyProfile: apiUpdateMyProfile,
  apiUpdateUserStatus: apiUpdateUserStatus,
  apiUploadLampiran: apiUploadLampiran,
  deleteFormReport: deleteFormReport,
  deleteFormTemplate: deleteFormTemplate,
  deleteIndikator: deleteIndikator,
  deleteProfilItem: deleteProfilItem,
  deleteUnit: deleteUnit,
  getCapaianTahunan: getCapaianTahunan,
  getFormReport: getFormReport,
  getFormTemplates: getFormTemplates,
  getMasterIndikator: getMasterIndikator,
  getMasterUnits: getMasterUnits,
  getProfilData: getProfilData,
  getRekapBulananUnit: getRekapBulananUnit,
  listBackupDB: listBackupDB,
  listFormReports: listFormReports,
  renumberIndikatorIds: renumberIndikatorIds,
  restoreDatabase: restoreDatabase,
  saveFormReport: saveFormReport,
  saveFormTemplate: saveFormTemplate,
  saveIndikator: saveIndikator,
  saveProfilItem: saveProfilItem,
  saveUnit: saveUnit,
  setValidasiForm: setValidasiForm,
  simpanDatabase: simpanDatabase
};

function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fn = String(req.fn || '');
    if (!Object.prototype.hasOwnProperty.call(API_MAP, fn)) throw new Error('Fungsi tidak tersedia: ' + fn);
    const res = API_MAP[fn].apply(null, Array.isArray(req.args) ? req.args : []);
    out = { ok: true, result: res === undefined ? null : res };
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
