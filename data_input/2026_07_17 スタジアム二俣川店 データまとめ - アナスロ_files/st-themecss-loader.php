@charset "UTF-8";










	.slick-prev,
	.slick-next,
	.slick-prev:hover,
	.slick-prev:focus,
	.slick-next:hover,
	.slick-next:focus {
		background-color: #424242	}







	.st-middle-menu .menu > li {
		width:33.33%;
	}
	.st-middle-menu .menu > li:nth-child(3n) {
		border-right:none;
	}
	.st-middle-menu .menu li a{
		font-size:80%;
	}


	.sns {
		width: 100%;
		text-align:center;
	}

	.sns li a {
		margin:0;
	}

	.post .sns ul,
	.sns ul {
		margin:0 auto;
		width:290px;
	}

	.snstext{
		display:none;
	}

	.snscount{
		display:none;
	}

	.sns li {
		float: left;
		list-style: none;
		width: 40px;
		margin-right: 10px;
		position: relative;
	}

	.sns li:last-child {
		margin-right: 0px;
	}

	.sns li i {
		font-size: 19px!important;
	}

	.sns li a {
					border-radius: 0;
				box-sizing: border-box;
		color: #fff;
		font-size: 19px;
		height: 40px;
		width: 40px;
		padding: 0;
		-webkit-box-pack: center;
		-webkit-justify-content: center;
		-ms-flex-pack: center;
		justify-content: center;
	}

	.sns li a .fa {
		padding: 0;
		border: 0;
		height: auto;
	}

	/* ツイッター */
	.sns .twitter a {
		box-shadow: none;
	}

	.sns .twitter a:hover {
		background:#4892cb;
		box-shadow: none;
	}

	.sns .fa-twitter::before {
		position: relative;
		top:1px;
		left: 1px;
	}

	/* Facebook */
	.sns .facebook a {
		box-shadow: none;
	}
	.sns .facebook a:hover {
		background:#2c4373;
		box-shadow: none;
	}

	.sns .fa-facebook::before {
		position: relative;
		top:1px;
	}

	/* グーグル */
	.sns .googleplus a {
		box-shadow: none;
	}
	.sns .googleplus a:hover {
		background:#d51e31;
		box-shadow: none;
	}

	.sns .fa-google-plus::before {
		position: relative;
		left: 1px;
	}

	/* URLコピー */
	.sns .share-copy a {
		box-shadow: none;
	}
	.sns .share-copy a:hover {
		background:#ccc;
		box-shadow: none;
	}

	.sns .share-copy::before {
		position: relative;
		left: 1px;
	}

	.sns .share-copy .fa-clipboard {
		border-right: none;
	}

	/* はてぶ */
	.sns .hatebu a {
		box-shadow: none;
	}

	.sns .hatebu a:hover {
		box-shadow: none;
		background:#00a5de;
	}

	.sns .st-svg-hateb::before {
		border-right:none;
		padding-right:0;
		font-size:19px!important;
	}

	.sns .st-svg-hateb::before {
		position: relative;
		left: 1px;
	}

	/* LINE */
	.sns .line a {
		box-shadow: none;
	}
	.sns .line a:hover {
		background:#219900;
		box-shadow: none;
	}

	.sns .fa-comment::before {
		position: relative;
		left: 1px;
		top: -1px;
	}

	/* Pocket */
	.sns .pocket a {
		box-shadow: none;
	}
	.sns .pocket a:hover {
		background:#F27985;
		box-shadow: none;
	}

	.sns .fa-get-pocket::before {
		position: relative;
		top: 1px;
	}

	/* アドセンス */
	.adbox,
	.adbox div {
		padding: 0!important;
	}






	.p-navi {
		display:none;
	}


	#headbox {
		padding: 10px!important;
	}

										.st-step-title, /* ステップ */
	.st-point .st-point-text, /* ポイント */
	.n-entry, /* NEW ENTRY */
	h4:not(.st-css-no) .point-in, /* 関連記事 */
	.cat-itiran p.point,
	.form-submit, /*コメント欄見出し*/
	.news-ca, /*お知らせタイトル*/
	.st-widgets-title, /* ウィジェットタイトル */
	.st-widgets-title span, /* ウィジェットタイトル */
	h4.menu_underh2 span,
	.st-header-flextitle,
    .post h2:not(.st-css-no),
    .post h2:not(.st-css-no) span,
    .post h3:not(.st-css-no),
    .post h3:not(.st-css-no) span,
	.h2modoki,
	.h3modoki,
    .entry-title:not(.st-css-no),
    .post .entry-title:not(.st-css-no) {
        font-family: 'Noto Sans JP', sans-serif;
    }
@media only screen and (max-width: 599px) {
	.st-header-flextitle {
		font-weight:900;
    }
}

	.st-cardbox .clearfix dd h5:not(.st-css-no),
	.post .st-cardbox .clearfix dd h5:not(.st-css-no),
	#side .st-cardbox .clearfix dd h5:not(.st-css-no) {
		border-bottom:none;
	}


	/*サイドバーカテゴリ*/
	#side li.cat-item a::after {
		content: " \f105";
		font-family: FontAwesome;
		position: absolute;
		right: 10px;
	}

	#side li.cat-item a {
		position: relative;
		vertical-align: middle;
		width:100%;
		padding: 10px;
		box-sizing:border-box;
		border-bottom: 1px solid #e1e1e1;
		color:#1a1a1a;
		text-decoration:none;
		display:block;
	}

	#side li.cat-item a:hover {
		opacity:0.5;
	}

			#side li.cat-item a {
			border-bottom-color: ;
			color: ;
		}
	



	/*スライドメニュー追加ボタン2*/
	#s-navi dt.trigger .acordion_extra_2 {
		max-width: 80%;
	}


/*
旧st-kanri.phpより移動（ここまで）
-------------------------------------------*/

/*グループ1
------------------------------------------------------------*/



/*縦一行目のセル*/
table tr td:first-child {
			}

/*横一行目のセル*/
table tr:first-child {
			}

/* 会話レイアウト */
	.st-kaiwa-hukidashi,
	.st-kaiwa-hukidashi2 {
		background-color: #FAFAFA;
	}
	.st-kaiwa-hukidashi:after {
		border-color: transparent #FAFAFA transparent transparent;
	}
	.st-kaiwa-hukidashi2:after {
		border-color: transparent transparent transparent #FAFAFA;
	}

/*この記事を書いた人*/
#st-tab-menu li.active {
  background: #212121;
}
#st-tab-box {
	border-color: #212121;
}
.post #st-tab-box p.st-author-post {
	border-bottom-color: #212121;
}
.st-author-date{
	color:#212121;
}

#st-tab-box {
	background:#FAFAFA;
}

/*こんな方におすすめ*/







	.st-blackboard-title:before {
  		content: "\f0f6\00a0";
  		font-family: FontAwesome;
	}

/*目次（TOC+）*/
#st_toc_container,
#toc_container {
			}

#st_toc_container:not(.st_toc_contracted):not(.only-toc),
#toc_container:not(.contracted) { /* 表示状態 */
		padding:15px 20px;
}

#st_toc_container:not(.st_toc_contracted):not(.only-toc),
#toc_container:not(.contracted) { /* 表示状態 */
	}

#st_toc_container.st_toc_contracted,
#toc_container.contracted { /* 非表示状態 */
	}







	#st_toc_container .st_toc_title:before,
	#toc_container .toc_title:before {
  		content: "\f0f6\00a0";
  		font-family: FontAwesome;
	}




    #st_toc_container li li li,
    #toc_container li li li {
		text-indent:-0.8em;
		padding-left:1em;
	}


/*マル数字olタグ*/


.post .maruno ol li:before {
			background: #212121;
				color:#ffffff;
	}

/*チェックulタグ*/


.post .maruck ul li:before {
			background: #212121;
				color:#ffffff;
	}

/*Webアイコン*/




	.st-blackboard.square-checkbox ul.st-blackboard-list:not(.st-css-no) li:before,
	.st-blackboard.square-checkbox ul li:after,
	.st-square-checkbox ul li:before,
	.st-square-checkbox ul li:after {	
		font-size: 150%;
	}

	.post .attentionmark2.on-color:not(.st-css-no):before,
	.post .fa-exclamation-triangle:not(.st-css-no) {
		color: #f44336;
}





/*サイト上部のボーダー色*/

/*ヘッダーの背景色*/
	
		#headbox-bg {
							background-color: transparent;
									background: none;
									}

		
	


/* header */
	header {
		background-image: url("https://ana-slo.com/wp-content/uploads/2022/08/header_pc_1400x250_50.jpg");
		background-position: center center;
        			background-repeat: no-repeat;
		        			background-size: cover;
		    }

/*ヘッダー下からの背景色*/
#content-w {
      
	}


input, textarea {
	color: #000;
}




	/*メインコンテンツの背景色*/
	main {
		background: #ffffff!important;
	}



	header .sitename a, /*ブログタイトル*/
	nav li a /* メニュー */
	{
		color: #ffffff;
	}

	/*ページトップ*/
	#page-top a {
		background: #212121;
	}

	#page-top a {
		line-height:100%;
		border-radius: 50%;
	}
	#page-top {
		right: 10px;
	}


/*キャプション */
	header h1,
	header .descr{
		color: #ffffff;
	}

	/* アコーディオン */
	#s-navi dt.trigger .op {
							color: #ffffff;
			}

/*モバイル用タイトルテキスト*/
	#st-mobile-logo a {
		color: #ffffff;
	}




/*アコーディオンメニュー内背景色*/
#s-navi dd.acordion_tree {
			background: #FAFAFA;
		/* 背景画像 */
	}

/*追加ボタン1*/
#s-navi dt.trigger .op-st {
			background: #424242;
				color: #ffffff;
	}

/*追加ボタン2*/
#s-navi dt.trigger .op-st2 {
			background: #424242;
				color: #ffffff;
	}



/*スマホフッターメニュー*/
#st-footermenubox a {
	color: #ffffff;
}

	#st-footermenubox {
  		background: #424242;
	}


	.acordion_tree ul.menu li a {
		border-bottom: 1px solid #424242;
	}

/* ガイドマップメニュー */





	.acordion_tree .st-ac-box ul.st-ac-cat {
	  border-top-color: #424242;
	  border-left-color: #424242;
	}

	.acordion_tree .st-ac-box ul.st-ac-cat > li.cat-item  {
	  border-right-color: #424242;
	  border-bottom-color: #424242;
	}



/*グループ2
------------------------------------------------------------*/
/* 投稿日時・ぱんくず・タグ */
#breadcrumb,
#breadcrumb div a,
div#breadcrumb a,
.blogbox p,
.tagst,
#breadcrumb ol li a,
#breadcrumb ol li h1,
#breadcrumb ol li,
.kanren:not(.st-cardbox) .clearfix dd .blog_info p,
.kanren:not(.st-cardbox) .clearfix dd .blog_info p a
{
	color: #616161;
}

/* 記事タイトル */




	
           	.post .entry-title:not(.st-css-no) {
				                                    background-color: transparent;
                                            background: none;
                                    
                                    border: none;
                
                
				
				            }

            
        


/* h2 */




	.h2modoki,
	.post h2:not(.st-css-no) {
		position: relative;
		padding-left:0;
		padding-bottom: 10px;
		border-top:none;
		
					padding-top:10px!important;
			padding-bottom:10px!important;
		                color: #000000;
                background-color:transparent;
	}
	.h2modoki::after,
	.post h2:not(.st-css-no)::after {
		position: absolute;
		bottom: -3px;
		left: 0;
		z-index: 2;
		content: '';
		width: 100%;
		height: 3px;
					/*Other Browser*/
			background: #0a0a0a;
			/* Android4.1 - 4.3 */
			background: -webkit-linear-gradient(left,  #bfbfbf 0%,#0a0a0a 100%);

			/* IE10+, FF16+, Chrome26+ */
			background: linear-gradient(to left,  #bfbfbf 0%,#0a0a0a 100%);
			}

	


/* h3 */




				.h3modoki,
            .post h3:not(.st-css-no):not(.st-matome):not(.rankh3):not(#reply-title) {
                background: #212121;
                color: #ffffff;
                position: relative;
                border: none;
                margin-bottom:30px;
		
					padding-top:10px!important;
			padding-bottom:10px!important;
		            }
        	.h3modoki:after,
            .post h3:not(.st-css-no):not(.st-matome):not(.rankh3):not(#reply-title):after {
                border-top: 10px solid #212121;
                content: '';
                position: absolute;
                border-right: 10px solid transparent;
                border-left: 10px solid transparent;
                bottom: -10px;
                left: 30px;
                border-radius: 2px;
            }
        	.h3modoki:before,
            .post h3:not(.st-css-no):not(.st-matome):not(.rankh3):not(#reply-title):before {
                border-top: 10px solid #212121;
                content: '';
                position: absolute;
                border-right: 10px solid transparent;
                border-left: 10px solid transparent;
                bottom: -10px;
                left: 30px;
            }

       		

	


/*h4*/


	.h4modoki,
    .post h4:not(.st-css-no):not(.st-matome):not(.rankh4):not(.point) {
                    border-left: 5px solid #000000;
                			color: #000000;
		                    background-color: #ffffff;
        
        
        
                    padding-left:20px;
        
                    padding-top:10px;
            padding-bottom:10px;
        
            }



/*まとめ*/

	/* 角丸 */
	.post .st-matome:not(.st-css-no):not(.rankh4):not(#reply-title):not(.point){
		border-radius:5px;
	}

	.post .st-matome:not(.st-css-no):not(.rankh4):not(#reply-title):not(.point){
		background: #adadad;
        			color: #000000;
				position: relative;
		border: none;
		margin-bottom:30px;
                    padding-left:20px!important;
        
                    padding-top:10px!important;
            padding-bottom:10px!important;
        
        	}

	.post .st-matome:not(.st-css-no):not(.rankh4):not(#reply-title):not(.point):after {
		border-top: 10px solid #adadad;
		content: '';
		position: absolute;
		border-right: 10px solid transparent;
		border-left: 10px solid transparent;
		bottom: -10px;
		left: 30px;
		border-radius: 2px;
	}

	.post .st-matome:not(.st-css-no):not(.rankh4):not(#reply-title):not(.point):before {
		border-top: 10px solid #adadad;
		content: '';
		position: absolute;
		border-right: 10px solid transparent;
		border-left: 10px solid transparent;
		bottom: -10px;
		left: 30px;
	}


/* ウィジェットタイトル */


	#side .menu_underh2,
    #side .st-widgets-title:not(.st-css-no) {
		font-weight:bold;
		margin-bottom: 10px;
	}



	#side .menu_underh2,
     #side .st-widgets-title:not(.st-css-no)  {
		position: relative;
		padding-left:0;
		padding-bottom: 10px;
		border-top:none;
					padding-left:15px!important;
		
					padding-top:5px!important;
			padding-bottom:5px!important;
		                color: #000000;
                background-color:transparent;
	}

	#side .menu_underh2::after,
    #side .st-widgets-title:not(.st-css-no) ::after {
		position: absolute;
		bottom: -3px;
		left: 0;
		z-index: 2;
		content: '';
		width: 100%;
		height: 3px;
					/*Other Browser*/
			background: #000000;
			/* Android4.1 - 4.3 */
			background: -webkit-linear-gradient(left,  #adadad 0%,#000000 100%);

			/* IE10+, FF16+, Chrome26+ */
			background: linear-gradient(to left,  #adadad 0%,#000000 100%);
			}

	

/* h5 */



	.h5modoki,
    .post h5:not(.st-css-no):not(.st-matome):not(.rankh5):not(.point):not(.st-cardbox-t):not(.popular-t):not(.kanren-t):not(.popular-t):not(.post-card-title) {
					border-left: 5px solid ;
		        			color: #000000;
							background-color: transparent;
		
		
		
					padding-left:15px!important;
		
					padding-top:7px!important;
			padding-bottom:7px!important;
		
			}



	.tagcloud a {
		border-color: #adadad;
		color: #adadad;
	}

	.post h4:not(.st-css-no):not(.rankh4).point,
	.cat-itiran p.point,
	.n-entry-t {
		border-bottom-color: #565656;
	}

	.post h4:not(.st-css-no):not(.rankh4) .point-in,
	.cat-itiran p.point .point-in,
	.n-entry {
			background-color: #565656;
				color: #ffffff;
		}

	.catname {
					background: #686868;
							color:#ffffff;
			}


	.post .st-catgroup a {
		color: #ffffff;
	}


/*グループ4
------------------------------------------------------------*/

.rssbox a {
	background-color: #212121;
}




/*ステップ
------------------------------------------------------------*/
.st-step {
	 		color: #ffffff;
				background: #212121;
				border-radius:5px;
	}
	.st-step:before{
		border-top-color: #212121;
	}

.st-step-title {
			  		border-bottom:solid 2px #212121;
	}

/* ポイント
------------------------------------------------------------*/
	.st-point:before {
					color: #ffffff;
							background: #212121;
			}

/*ブログカード
------------------------------------------------------------*/
/* 枠線 */

/* ラベル */
.st-cardbox-label-text {
		}

/*フリーボックスウィジェット
------------------------------------------------------------*/
.freebox {
			border-top-color: #212121;
				background: #FAFAFA;
	}

.p-entry-f {
			background: #212121;
				color: #ffffff;
	}

/* エリア内テキスト */

/*メモボックス
------------------------------------------------------------*/

/*スライドボックス
------------------------------------------------------------*/

/*お知らせ
------------------------------------------------------------*/
/*お知らせバーの背景色*/
#topnews-box div.rss-bar {
			border-color: #212121;
	
			/*Other Browser*/
		background: #212121;
		/*For Old WebKit*/
		background: -webkit-linear-gradient( #212121 0%, #212121 100% );
		/*For Modern Browser*/
		background: linear-gradient( #212121 0%, #212121 100% );
	
			color: #ffffff;
	}

/*お知らせ日付の文字色*/
#topnews-box dt {
	color: #212121;
}

#topnews-box div dl dd a {
	color: #000000;
}

#topnews-box dd {
	border-bottom-color: #212121;
}

#topnews-box {
			background-color:transparent!important;
	}


/*追加カラー
------------------------------------------------------------*/

#footer .footerlogo a, /* フッターロゴ */
#footer .footerlogo,
#footer .footer-description a, /* フッター説明 */
#footer .footer-description,
#footer .head-telno a, /* フッターの電話番号 */
#footer .head-telno, /* フッターの電話番号 */
#footer .widget_recent_entries ul li a, /* 最近の投稿 */
#footer .widget_recent_entries ul li,
#footer .widget_categories ul li a, /* カテゴリ（デフォルト） */
#footer .widget_categories ul li,
#footer .rankh3:not(.st-css-no),
#footer .st_side_rankwidgets a, /* ランキングウィジェットタイトル */
#footer .rankwidgets-cont p, /* 説明 */
#footer .kanren dd a, /* 記事一覧タイトル */
#footer .kanren .st-excerpt p, /*  記事一覧説明 */
#footer .kanren .blog_info p, /*  記事一覧日付 */
footer .footer-l *,
footer .footer-l a,
#footer .menu-item a,
#footer .copyr,
#footer .copyr a,
#footer .copy,
#footer .copy a {
	color: #ffffff !important;
}

#footer #newsin dt, /* お知らせ日時 */
#footer #newsin dd a, /* お知らせ日テキスト*/
#footer #newsin dd, /* お知らせ日ボーダー */
#footer .cat-item a,
#footer .tagcloud a {
			color: #ffffff;	
		border-color: #ffffff;
	}

#footer .st-widgets-title {
			color: #ffffff;	
		font-weight: bold;
}

	footer .footermenust li {
		border-right-color: #ffffff !important;
	}

/*フッター背景色*/

	#footer {
        max-width:100%;
					/*Other Browser*/
			background: #212121;
			/* Android4.1 - 4.3 */
			background: url("https://ana-slo.com/wp-content/uploads/2022/08/header_pc_1400x250_50.jpg"), -webkit-linear-gradient(left,  #424242 0%,#212121 100%);

			/* IE10+, FF16+, Chrome26+ */
			background: url("https://ana-slo.com/wp-content/uploads/2022/08/header_pc_1400x250_50.jpg"), linear-gradient(to left,  #424242 0%,#212121 100%);

				}

			#footer {
			background-position: center center;
							background-repeat: no-repeat;
										background-size: cover;
						}
	

/*任意の人気記事
------------------------------------------------------------*/
	.st-pvm-ranking-item-image::before, /* PVモニター */
	.poprank-no {
		background: #212121;
	}

.post .p-entry, #side .p-entry, .home-post .p-entry {
			background: #212121;
				color: #ffffff;
	}

.pop-box, .nowhits .pop-box, .nowhits-eye .pop-box,
.st-eyecatch + .nowhits .pop-box {
			border-top-color: #212121;
				background: #FAFAFA;
	}

	.p-entry::after {
		border-bottom: 5px solid #212121;
		border-left: 5px solid #212121;
	}

.pop-box:not(.st-wpp-views-widgets),
#side aside .kanren.pop-box:not(.st-wpp-views-widgets) {
			padding:20px 20px 10px;
	}

.pop-box:not(.st-wpp-views-widgets),
#side aside .kanren.pop-box:not(.st-wpp-views-widgets) {
		padding:30px 20px 10px;
		border: none;
}


	.poprank-no2,
	.poprank-no {
			background: #212121;
				color: #ffffff!important;
		}

/*WordPressPopularPosts連携*/

#st-magazine .st-wp-views, /*CARDs JET*/
#st-magazine .st-wp-views-limit, /*CARDs JET*/
.st-wppviews-label .wpp-views, /*Ex*/
.st-wppviews-label .wpp-views-limit, /*Ex*/
.st-wpp-views-widgets .st-wppviews-label .wpp-views {
			color: #ffffff;
				background:#212121;
	}

/*ウィジェット問合せボタン*/

.st-formbtn {
			border-radius: 3px;
	
	
			/*Other Browser*/
		background: #212121;
	}

	.st-formbtn .st-originalbtn-r {
		border-left-color: #ffffff;
	}

	a.st-formbtnlink {
		color: #ffffff;
	}

/*ウィジェットオリジナルボタン*/

.st-originalbtn {
			border-radius: 3px;
	
	
			/*Other Browser*/
		background: #212121;
	}

	.st-originalbtn .st-originalbtn-r {
		border-left-color: #ffffff;
	}

	a.st-originallink {
		color: #ffffff;
	}

/*ミドルメニュー
------------------------------------------------------------*/
	.st-middle-menu {
					background-color: #2e2e2e;
				}

.st-middle-menu ul{
	}

.st-middle-menu .menu > li{
	}

.st-middle-menu .menu > li:last-child {
		border-bottom: none;
}

	
.st-middle-menu .menu li a{
			color: #ffffff;
	}



/*サイドメニューウィジェット
------------------------------------------------------------*/
/*背景色*/
#sidebg {
			background: #FAFAFA;
		}

/*liタグの階層*/
#side aside .st-pagelists ul li:not(.sub-menu) {
			border: none;
	}

#side aside .st-pagelists ul .sub-menu li {
	border: none;
}

#side aside .st-pagelists ul li:last-child {
			border-bottom: none;
	}

#side aside .st-pagelists ul .sub-menu li:first-child {
			border-top: none;
	}

#side aside .st-pagelists ul li li:last-child {
	border: none;
}

#side aside .st-pagelists ul .sub-menu .sub-menu li {
	border: none;
}


#side aside .st-pagelists ul li a {
			color: #ffffff;
				/*Other Browser*/
		background: #212121;
		/* Android4.1 - 4.3 */
		background: url(""), -webkit-linear-gradient(left,  #424242 0%,#212121 100%);

		/* IE10+, FF16+, Chrome26+ */
		background: url(""), linear-gradient(to left,  #424242 0%,#212121 100%);

	}




#side aside .st-pagelists .sub-menu a {
			border-bottom-color: #424242;
		color: #212121;
}

#side aside .st-pagelists .sub-menu .sub-menu li:last-child {
	border-bottom: 1px solid #424242;
}

#side aside .st-pagelists .sub-menu li .sub-menu a,
#side aside .st-pagelists .sub-menu li .sub-menu .sub-menu li a {
	color: #212121;
}

	#side aside .st-pagelists .sub-menu li .sub-menu a:hover,
	#side aside .st-pagelists .sub-menu li .sub-menu .sub-menu li a:hover,
	#side aside .st-pagelists .sub-menu a:hover {
		opacity:0.8;
		color: #212121;
	}


	#side aside .st-pagelists ul li a {
		padding-left:15px;
	}

	#side aside .st-pagelists ul li a {
		padding-top:8px;
		padding-bottom:8px;
	}

/*Webアイコン*/
	#side aside .st-pagelists ul li a:before {
					content: "\f138\00a0\00a0";
				font-family: FontAwesome;
									color:#ffffff;
					
	}
	#side aside .st-pagelists li li a:before {
		content: none;
	}

	#side aside .st-pagelists li li a:before {
		content: "\f105\00a0\00a0";
		font-family: FontAwesome;
					color:#212121;
			}

.wpcf7-submit {
			background: #212121;
				color: #ffffff;
	}

/* ヘッダー画像エリア */

/* メイン画像背景色 */


/*強制センタリング・中央寄せ
------------------------------------------------------------*/
	    
            .entry-content .h2modoki,
        .entry-content h2:not(.st-css-no)
        {
            text-align:center;
								padding-left:10px;
					padding-right:10px;
			        }
                
            .entry-content .h3modoki,
        .entry-content h3:not(.st-css-no):not(.st-matome):not(.rankh3):not(#reply-title)
        {
            text-align:center;
								padding-left:10px;
					padding-right:10px;
			        }
                    .entry-content .h3modoki:after,
            .entry-content .h3modoki:before,
            .entry-content h3:not(.st-css-no):not(.st-matome):not(.rankh3):not(#reply-title):after,
            .entry-content h3:not(.st-css-no):not(.st-matome):not(.rankh3):not(#reply-title):before {
                left: calc(50% - 10px);
            }
                
/*media Queries タブレットサイズ（959px以下）
----------------------------------------------------*/
@media only screen and (max-width: 959px) {

	/*-------------------------------------------
	旧st-kanri.phpより移動（ここから）
	*/

	
			
	/*
	旧st-kanri.phpより移動（ここまで）
	-------------------------------------------*/

	/*-- ここまで --*/
}

/*media Queries タブレットサイズ以上
----------------------------------------------------*/
@media only screen and (min-width: 600px) {

    /*-------------------------------------------
    旧st-kanri.phpより移動（ここから）
    */
    
        
        
        
    
	
	    
        
        
    /*
    旧st-kanri.phpより移動（ここまで）
    -------------------------------------------*/
	/* 目次 */
	#st_toc_container:not(.st_toc_contracted):not(.only-toc),
	#toc_container:not(.contracted) { /* 表示状態 */
			}

	/*強制センタリング・中央寄せ
	------------------------------------------------------------*/
				
				
				
}

/*media Queries タブレットサイズ（600px～959px）のみで適応したいCSS -タブレットのみ
---------------------------------------------------------------------------------------------------*/
@media only screen and (min-width: 600px) and (max-width: 959px) {

	/*-------------------------------------------
	旧st-kanri.phpより移動（ここから）
	*/

    /*--------------------------------
    各フォント設定
    ---------------------------------*/
    
    
    /*基本のフォントサイズ*/
    .post .entry-content p:not(.p-entry-t):not(.p-free):not(.sitename):not(.post-slide-title):not(.post-slide-date):not(.post-slide-more):not(.st-catgroup):not(.wp-caption-text):not(.cardbox-more):not(.st-minihukidashi):not(.st-mybox-title):not(.st-memobox-title):not(.st-mybtn), /* テキスト */
    .post .entry-content .st-kaiwa-hukidashi, /* ふきだし */
    .post .entry-content .st-kaiwa-hukidashi2, /* ふきだし */
    .post .entry-content .yellowbox, /* 黄色ボックス */
    .post .entry-content .graybox, /* グレーボックス */
    .post .entry-content .redbox, /* 薄赤ボックス */
    .post .entry-content #topnews .clearfix dd p, /* 一覧文字 */
    .post .entry-content ul li, /* ulリスト */
    .post .entry-content ol li, /* olリスト */
	.post .entry-content #st_toc_container > ul > li, /* 目次用 */
    .post .entry-content #comments #respond, /* コメント */
    .post .entry-content #comments h4, /* コメントタイトル */
	.post .entry-content h5:not(.kanren-t):not(.popular-t):not(.st-cardbox-t), /* H5 */
	.post .entry-content h6 {
					}

	/* スライドの抜粋 */
	.post .entry-content .post-slide-excerpt p:not(.p-entry-t):not(.p-free):not(.sitename):not(.post-slide-title):not(.post-slide-date):not(.post-slide-more):not(.st-catgroup):not(.wp-caption-text):not(.cardbox-more):not(.st-minihukidashi):not(.st-mybox-title):not(.st-memobox-title):not(.st-mybtn),
	.post .entry-content .st-excerpt p:not(.p-entry-t):not(.p-free):not(.sitename):not(.post-slide-title):not(.post-slide-date):not(.post-slide-more):not(.st-catgroup):not(.wp-caption-text):not(.cardbox-more):not(.st-minihukidashi):not(.st-mybox-title):not(.st-memobox-title):not(.st-mybtn),
	.post .entry-content .st-card-excerpt p:not(.p-entry-t):not(.p-free):not(.sitename):not(.post-slide-title):not(.post-slide-date):not(.post-slide-more):not(.st-catgroup):not(.wp-caption-text):not(.cardbox-more):not(.st-minihukidashi):not(.st-mybox-title):not(.st-memobox-title):not(.st-mybtn),
	.post .entry-content .kanren:not(.st-cardbox) .clearfix dd p:not(.p-entry-t):not(.p-free):not(.sitename):not(.post-slide-title):not(.post-slide-date):not(.post-slide-more):not(.st-catgroup):not(.wp-caption-text):not(.cardbox-more):not(.st-minihukidashi):not(.st-mybox-title):not(.st-memobox-title):not(.st-mybtn){
					}

	
    /* 記事タイトル */
    #contentInner .post .entry-title:not(.st-css-no2) {
				    }
    
    /* H2 */
    .post .entry-content h2:not(.st-css-no2),
    .post .entry-content .h2modoki{
				    }
    
    /* H3 */
    .post .entry-content h3:not(.st-css-no2):not(#reply-title),
    .post .entry-content .h3modoki {
				    }
    
    /* H4 */
    .post .entry-content h4:not(.st-css-no2):not(.point),
    .post .entry-content .h4modoki {
				    }

	/*
	旧st-kanri.phpより移動（ここまで）
	-------------------------------------------*/

	
/*-- ここまで --*/
}


/*media Queries PCサイズ
----------------------------------------------------*/
@media only screen and (min-width: 960px) {

	/*-------------------------------------------
	旧st-kanri.phpより移動（ここから）
	*/

    /*--------------------------------
    各フォント設定
    ---------------------------------*/
    
        
    /*基本のフォントサイズ*/
    .post .entry-content p:not(.p-entry-t):not(.p-free):not(.sitename):not(.post-slide-title):not(.post-slide-date):not(.post-slide-more):not(.st-catgroup):not(.wp-caption-text):not(.cardbox-more):not(.st-minihukidashi):not(.st-mybox-title):not(.st-memobox-title):not(.st-mybtn), /* テキスト */
    .post .entry-content .st-kaiwa-hukidashi, /* ふきだし */
    .post .entry-content .st-kaiwa-hukidashi2, /* ふきだし */
    .post .entry-content .yellowbox, /* 黄色ボックス */
    .post .entry-content .graybox, /* グレーボックス */
    .post .entry-content .redbox, /* 薄赤ボックス */
    .post .entry-content #topnews .clearfix dd p, /* 一覧文字 */
    .post .entry-content ul li, /* ulリスト */
    .post .entry-content ol li, /* olリスト */
	.post .entry-content #st_toc_container > ul > li, /* 目次用 */
    .post .entry-content #comments #respond, /* コメント */
    .post .entry-content #comments h4, /* コメントタイトル */
	.post .entry-content h5:not(.kanren-t):not(.popular-t):not(.st-cardbox-t), /* H5 */
	.post .entry-content h6 {
					}

	/* スライドの抜粋 */
	.post .entry-content .post-slide-excerpt p:not(.p-entry-t):not(.p-free):not(.sitename):not(.post-slide-title):not(.post-slide-date):not(.post-slide-more):not(.st-catgroup):not(.wp-caption-text):not(.cardbox-more):not(.st-minihukidashi):not(.st-mybox-title):not(.st-memobox-title):not(.st-mybtn),
	.post .entry-content .st-excerpt p:not(.p-entry-t):not(.p-free):not(.sitename):not(.post-slide-title):not(.post-slide-date):not(.post-slide-more):not(.st-catgroup):not(.wp-caption-text):not(.cardbox-more):not(.st-minihukidashi):not(.st-mybox-title):not(.st-memobox-title):not(.st-mybtn),
	.post .entry-content .st-card-excerpt p:not(.p-entry-t):not(.p-free):not(.sitename):not(.post-slide-title):not(.post-slide-date):not(.post-slide-more):not(.st-catgroup):not(.wp-caption-text):not(.cardbox-more):not(.st-minihukidashi):not(.st-mybox-title):not(.st-memobox-title):not(.st-mybtn),
	.post .entry-content .kanren:not(.st-cardbox) .clearfix dd p:not(.p-entry-t):not(.p-free):not(.sitename):not(.post-slide-title):not(.post-slide-date):not(.post-slide-more):not(.st-catgroup):not(.wp-caption-text):not(.cardbox-more):not(.st-minihukidashi):not(.st-mybox-title):not(.st-memobox-title):not(.st-mybtn){
					}

	
    /* 記事タイトル */
    #contentInner .post .entry-title:not(.st-css-no2) {
				    }
    
    /* H2 */
    .post .entry-content h2:not(.st-css-no2),
    .post .entry-content .h2modoki{
				    }
    
    /* H3 */
    .post .entry-content h3:not(.st-css-no2):not(#reply-title),
    .post .entry-content .h3modoki {
				    }
    
    /* H4 */
    .post .entry-content h4:not(.st-css-no2):not(.point),
    .post .entry-content .h4modoki {
				    }

	/*--------------------------------
	全体のサイズ
	---------------------------------*/

	#st-menuwide, /*メニュー*/
	nav.smanone,
	nav.st5,
	#st-header-cardlink,
	#st-menuwide div.menu,
	#st-menuwide nav.menu,
	#st-header, /*ヘッダー*/
	#st-header-under-widgets-box, /*ヘッダー画像下*/
	#content, /*コンテンツ*/
	#footer-in /*フッター*/
	 { 
		max-width:1114px;
	}

	.st-lp-wide #content /* LPワイド */
	 { 
		max-width:100%;
	}

	#headbox
	 { 
		max-width:1094px;
	}

	
	/*1カラムの幅のサイズ*/
	.colum1:not(.st-lp-wide) #st-header-under-widgets-box,
	.colum1:not(.st-lp-wide) #content {
    	max-width: 1114px;
	}

	/* ヘッダー画像/記事スライドショー横並び */
	
	    
        /*--------------------------------
        PCのレイアウト（右サイドバー）
        ---------------------------------*/
    
        #contentInner {
            float: left;
            width: 100%;
            margin-right: -300px;
        }
    
        main {
            margin-right: 320px;
            margin-left: 0px;
            background-color: #fff;
            border-radius: 4px;
            -webkit-border-radius: 4px;
            -moz-border-radius: 4px;
            padding: 30px 50px 30px;
        }
    
        #side aside {
            float: right;
            width: 300px;
            padding: 0px;
        }
    
    
        
    /**
     * サイト名とキャッチフレーズ有無の調整
     */
    
        
        
            #header-r .footermenust {
            margin: 0;
        }
        
            header .sitename {
                padding: 5px;
            margin: 0;
                            line-height:0;
                font-size:1px;
                    }
        #headbox {
            padding: 5px 10px!important;
        }
        
        
            /*PCアドセンスを横並び*/
        .adbox:after {
            content: "";
            display: block;
            clear: both;
        }
        .adbox div {
            float:left;
            margin-right:20px;
            padding-top:0!important;
            padding-bottom:10px;
        }
    
        .adbox div:last-child {
            margin-right:0px;
        }
        
        
    /*
    旧st-kanri.phpより移動（ここまで）
    -------------------------------------------*/

	/*TOC+*/
	#toc_container:not(.contracted) { /* 表示状態 */
					}

	/*ヘッダーの背景色*/
	
	/*メインコンテンツのボーダー*/
	
	
	/* メイン画像100% */
	
	/*wrapperに背景がある場合*/
	
	/*メニュー*/
	#st-menuwide {
			border-top: none;
		border-bottom: none;
				border-left: none;
		border-right: none;
	
	
			background-image: url("");		
		background-color: #2e2e2e;
		}

	

	header .smanone ul.menu li, 
	header nav.st5 ul.menu  li,
	header nav.st5 ul.menu  li,
	header #st-menuwide div.menu li,
	header #st-menuwide nav.menu li
	{
			border-right: none;
		}

	header .smanone ul.menu li li,
	header nav.st5 ul.menu li li,
	header #st-menuwide div.menu li li,
	header #st-menuwide nav.menu li li {
    	border:none;
	}

		header .smanone ul.menu li a, 
	header nav.st5 ul.menu  li a,
	header #st-menuwide div.menu li a,
	header #st-menuwide nav.menu li a,
	header .smanone ul.menu li a:hover, 
	header nav.st5 ul.menu  li a:hover,
	header #st-menuwide div.menu li a:hover,
	header #st-menuwide nav.menu li a:hover{
		color: #ffffff;
	}
	
	header .smanone ul.menu li:hover, 
	header nav.st5 ul.menu  li:hover,
	header #st-menuwide div.menu li:hover,
	header #st-menuwide nav.menu li:hover{
		background: rgba(255,255,255,0.1);
	}

	header .smanone ul.menu li li a:hover, 
	header nav.st5 ul.menu  li li a:hover,
	header #st-menuwide div.menu li li a:hover,
	header #st-menuwide nav.menu li li a:hover{
		opacity:0.9;
	}

	
			header .smanone ul.menu li li a {
					background: #424242;
							border-top-color: #2e2e2e;
				}
	
	/*メニューの上下のパディング*/
		

	/* グローバルメニュー */
			#st-menuwide {
			max-width: 100%;
					}
	
	
	/*ヘッダーウィジェット*/
	header .headbox .textwidget,
	#footer .headbox .textwidget{
					background: #FAFAFA;
					}

			/*ヘッダーの電話番号とリンク色*/
		.head-telno a, #header-r .footermenust a {
			color: #ffffff;
		}
	
			#header-r .footermenust li {
			border-right-color: #ffffff;
		}
	
			/*トップ用おすすめタイトル*/
		.nowhits .pop-box {
			border-top-color: #212121;
		}
	
	/*記事エリアを広げる*/
	
	/*強制センタリング・中央寄せ
	------------------------------------------------------------*/
				
				
				
			#side .smanone.st-excerpt {
			display:none;
		}
	
/*-- ここまで --*/
}

/*media Queries スマートフォンのみ（599px）以下
---------------------------------------------------------------------------------------------------*/
@media only screen and (max-width: 599px) {

    /*-------------------------------------------
    旧st-kanri.phpより移動（ここから）
    */
    
    /*--------------------------------
    各フォント設定
    ---------------------------------*/
    
        
    /*基本のフォントサイズ*/
    .post .entry-content p:not(.p-entry-t):not(.p-free):not(.sitename):not(.post-slide-title):not(.post-slide-date):not(.post-slide-more):not(.st-catgroup):not(.wp-caption-text):not(.cardbox-more):not(.st-minihukidashi):not(.st-mybox-title):not(.st-memobox-title):not(.st-mybtn), /* テキスト */
    .post .entry-content .st-kaiwa-hukidashi, /* ふきだし */
    .post .entry-content .st-kaiwa-hukidashi2, /* ふきだし */
    .post .entry-content .yellowbox, /* 黄色ボックス */
    .post .entry-content .graybox, /* グレーボックス */
    .post .entry-content .redbox, /* 薄赤ボックス */
    .post .entry-content #topnews .clearfix dd p, /* 一覧文字 */
    .post .entry-content ul li, /* ulリスト */
    .post .entry-content ol li, /* olリスト */
    .post .entry-content #st_toc_container > ul > li, /* 目次用 */
    .post .entry-content #comments #respond, /* コメント */
    .post .entry-content #comments h4, /* コメントタイトル */
    .post .entry-content h5:not(.kanren-t):not(.popular-t):not(.st-cardbox-t), /* H5 */
    .post .entry-content h6 {
                    }
    
    /* スライドの抜粋 */
    .post .entry-content .post-slide-excerpt p:not(.p-entry-t):not(.p-free):not(.sitename):not(.post-slide-title):not(.post-slide-date):not(.post-slide-more):not(.st-catgroup):not(.wp-caption-text):not(.cardbox-more):not(.st-minihukidashi):not(.st-mybox-title):not(.st-memobox-title):not(.st-mybtn),
    .post .entry-content .st-excerpt p:not(.p-entry-t):not(.p-free):not(.sitename):not(.post-slide-title):not(.post-slide-date):not(.post-slide-more):not(.st-catgroup):not(.wp-caption-text):not(.cardbox-more):not(.st-minihukidashi):not(.st-mybox-title):not(.st-memobox-title):not(.st-mybtn),
    .post .entry-content .st-card-excerpt p:not(.p-entry-t):not(.p-free):not(.sitename):not(.post-slide-title):not(.post-slide-date):not(.post-slide-more):not(.st-catgroup):not(.wp-caption-text):not(.cardbox-more):not(.st-minihukidashi):not(.st-mybox-title):not(.st-memobox-title):not(.st-mybtn),
    .post .entry-content .kanren:not(.st-cardbox) .clearfix dd p:not(.p-entry-t):not(.p-free):not(.sitename):not(.post-slide-title):not(.post-slide-date):not(.post-slide-more):not(.st-catgroup):not(.wp-caption-text):not(.cardbox-more):not(.st-minihukidashi):not(.st-mybox-title):not(.st-memobox-title):not(.st-mybtn){
                    }
    
        
    /* 記事タイトル */
    #contentInner .post .entry-title:not(.st-css-no2) {
                    }
    
    /* H2 */
    .post .entry-content h2:not(.st-css-no2),
    .post .entry-content .h2modoki{
                    }
    
    /* H3 */
    .post .entry-content h3:not(.st-css-no2):not(#reply-title),
    .post .entry-content .h3modoki {
                    }
    
    /* H4 */
    .post .entry-content h4:not(.st-css-no2):not(.point),
    .post .entry-content .h4modoki {
                    }
    
    /*
    旧st-kanri.phpより移動（ここまで）
    -------------------------------------------*/

	
/*-- ここまで --*/
}

/*-------------------------------------------
旧st-kanri.phpより移動（ここから）
*/


/*
旧st-kanri.phpより移動（ここまで）
-------------------------------------------*/

/*-------------------------------------------
ガイドマップメニュー（ここから）
*/
/*
ガイドマップメニュー（ここまで）
-------------------------------------------*/
