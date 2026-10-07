
# Excel vbaでe支店APIを利用するサンプル 

	ファイル名: e_api_Excel_pubkey.xlsm
	言語：Excel VBA
	APIバージョン： V4r10で動作確認


-----------------------------------------
	
	ご注意！！ 
	 本番環境に接続した場合、実際に注文を出します。
	 市場で約定した場合、取り消せません。
	 十分に注意してご利用ください。

-----------------------------------------
	
	このExcelでのAPIの利用は、認証情報を含むファイルが第三者に渡った場合に、
	認証情報を容易に復元できないよう考慮してあります。
	しかし万全ではありません。
	
	PC自体が不正アクセスを受けた場合等も含め、
	利用者ご自身で適切なセキュリティ対策を行ってください。
	
-----------------------------------------
	
1）動作テストを実行した環境

	os:  Windows 11 Pro 24H2
	excel: 2016 64ビット

２）Excelの設定等

	1.EXCELのメニューからファイル－＞オプション－＞リボンのユーザ設定で開発にチェックを入れる。
	結果メニューに「開発」が表示されます。
	
	2.開発－＞Visual Basic を選択、標準モジュールに上記「３．提供ファイル」、2,3 のモジュールを追加します。
	
	3.VBA のメニュー「ツールー＞参照設定」で以下２つをチェック（参照設定）します。
	  ・Microsoft XML, v6.0
	  ・Microsoft Scripting Runtime
	※EXCEL を新規にインストールした状態で動作確認をしておりますので、
	EXECELの各種設定等を変更されている場合は初期状態に戻しご利用下さい。

	また、VBA での JSON 文字列解析に 
		VBA-JSON v2.3.1（https://github.com/VBA-tools/VBA-JSON/releases/tag/v2.3.1）（MIT License）
	を利用しています。


３）APIの利用には、事前に立花証券ｅ支店に口座開設が必要です。


４）利用前に以下の手順書に従い準備してください。
	
	認証ID・秘密鍵等の取得方法.pdf
	認証情報の暗号化などの事前準備.pdf
	   
	
５）利用の手順

	・先ず「ログイン」実行してください。ログイン情報がログインシートに保存されます。
	・ログインで取得したセッションが有効な間、ログイン情報を利用して他のコマンドを実行できます。
	・各プログラムの実行には、
			肌色のセルに、必要項目を入力してください。
			青色のセルに、返信データを一部出力します。
	・シート「api_応答」に、APIへの送信テキストと返信データを出力します。
	・シート「結果」に、返信データを簡単に出力します。


６）実行内容は、以下になります。

	シート	オブジェクト名	プロシージャー名
	--- 	Login	func_login
	--- 	Logout	func_logout
	照会	照会_現物_買付可能額	get_CLMZanKaiKanougaku
	照会	照会_信用_新規建可能額	get_CLMZanShinkiKanoIjiritu
	照会	照会_注文約定一覧	get_CLMOrderList
	照会	照会_注文約定一覧_詳細	get_CLMOrderListDetail
	現物	現物_買い注文	order_gen_buy
	現物	現物_預り株一覧	get_CLMGenbutuKabuList
	現物	現物_売り注文	order_gen_sell
	信用	信用_新規_買い注文	order_shinki_buy
	信用	信用_新規_売り注文	order_shinki_sell
	信用	信用_建玉一覧	get_CLMShinyouTategyokuList
	信用	信用_返済_買い注文	order_hensai_buy
	信用	信用_返済_売り注文	order_hensai_sell
	信用	信用_返済_買い注文_個別指定	order_hensai_buy_aCLMKabuHensaiData
	信用	信用_返済_売り注文_個別指定	order_hensai_sell_aCLMKabuHensaiData
	訂正・取消 取消	cancel_order
	訂正・取消 一括取消	cancel_all_order
	訂正・取消 訂正	correct_order
	株価	日足データ取得		get_CLMMfdsGetMarketPriceHistory
	株価	スナップショット		get_CLMMfdsGetMarketPrice
	マスター	1.株式銘柄マスタ問合取得		get_master_1_CLMStkGetIssueMstKabu
	マスター	2.株式銘柄市場マスタ問合取得	get_master_2_CLMStkGetIssueSizyouMstKabu
	マスター	3.先物銘柄マスタ問合取得		get_master_3_CLMStkGetIssueMstSak
	マスター	4.オプション銘柄マスタ問合取得	get_master_4_CLMStkGetIssueMstOp
	マスター	5.指数銘柄マスタ問合取得		get_master_5_CLMStkGetIssueMstIndex
	マスター	6.為替銘柄マスタ問合取得		get_master_6_CLMStkGetIssueMstFx
	マスター	7.日付情報問合取得			get_master_7_CLMStkGetDateZyouhou
	マスター	8.呼値情報問合取得			get_master_8_CLMStkGetYobine
	マスター	9.代用掛目情報問合取得		get_master_9_CLMStkGetDaiyouKakeme
	マスター	10.株式銘柄別・市場別規制情報問合取得		get_master_10_CLMStkGetIssueSizyouKiseiKabu
	マスター	11.派生銘柄別・市場別規制情報問合取得		get_master_11_CLMStkGetIssueSizyouKiseiHasei
	マスター	12.保証金マスタ情報問合取得				get_master_12_CLMStkGetHosyoukinMst
	マスター	13.取引所エラー等理由コード情報問合取得	get_master_13_CLMStkGetOrderErrReason
	
 	（現在、マスターデータの更新に不具合が発生しており、誤った情報が記載されています。ご注意ください。）
	（株式 銘柄マスタ（CLMIssueMstKabu）では、銘柄コード、銘柄名、銘柄名略称、銘柄名（カナ）、銘柄名（英語表記）、優先市場、業種コード、業種コード名 のみ利用できます。その他のデータについてはｅ支店サポートセンターにご確認ください。）


７）利用時間外に接続した場合、"p_errno":"9"（システム、サービス停止中。）が返されます。詳しくは「立花証券・ｅ支店・ＡＰＩ、EVENT I/F 利用方法、データ仕様」4ページをご参照ください。デモ環境の利用時間は、デモ環境の説明ページを参照ください。

８）本サンプルプログラムは、事務方の者が休日や空き時間に作成したため、色々足りておりません。ご容赦ください。

９）本プログラムはライセンスの範囲内で自由にご利用ください。

１０）このソフトウェアを使用したことによって生じたすべての障害・損害・不具合等に関して、私と私の関係者および私の所属するいかなる団体・組織とも、一切の責任を負いません。各自の責任においてご使用ください。
