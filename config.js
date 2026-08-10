window.APP_CONFIG = Object.freeze({
  // Google Cloud Console 建立「Web application」OAuth Client 後填入。
  // 這是公開識別碼，不是 Client Secret。
  googleClientId: "758952394198-k481nt0rur39neujnh603798eobm9s7c.apps.googleusercontent.com",
  targetSheetName: "續訂清單",
  headers: Object.freeze({
    id: "編號細目",
    currentReply: "本月回覆",
    previousReply: "上個月回覆",
  }),
  previewLimit: 100,
});
