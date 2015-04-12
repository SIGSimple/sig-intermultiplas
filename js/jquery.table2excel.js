/*
 *  jQuery table2excel - v1.0.1
 *  jQuery plugin to export an .xls file in browser from an HTML table
 *  https://github.com/rainabba/jquery-table2excel
 *
 *  Made by rainabba
 *  Under MIT License
 */
//table2excel.js
;(function ( $, window, document, undefined ) {
		var pluginName = "table2excel",
				defaults = {
				exclude: ".noExl",
                name: "Table2Excel"
		};

		// The actual plugin constructor
		function Plugin ( element, options ) {
				this.element = element;
				// jQuery has an extend method which merges the contents of two or
				// more objects, storing the result in the first object. The first object
				// is generally empty as we don't want to alter the default options for
				// future instances of the plugin
				this.settings = $.extend( {}, defaults, options );
				this._defaults = defaults;
				this._name = pluginName;
				this.init();
		}

		Plugin.prototype = {
			init: function () {
				var e = this;
				e.template = "<html xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\" xmlns=\"http://www.w3.org/TR/REC-html40\"><head><!--[if gte mso 9]><xml>";
				e.template += "<x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions>";
				e.template += "<x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body><table>{table}</table></body></html>";
				e.tableRows = "";

				// get contents of table except for exclude
				$(e.element).find("tr").not(this.settings.exclude).each(function (i,o) {
					e.tableRows += "<tr>" + $(o).html() + "</tr>";
				});
				this.tableToExcel(this.tableRows, this.settings.name);
			},
			tableToExcel: function (table, name) {
				var e = this;
				
				e.uri = "data:application/vnd.ms-excel;base64,";
				
				e.base64 = function (s) {
					return window.btoa(unescape(encodeURIComponent(s)));
				};

				e.base64_encode = function(data) {
					//  discuss at: http://phpjs.org/functions/base64_encode/
					// original by: Tyler Akins (http://rumkin.com)
					// improved by: Bayron Guevara
					// improved by: Thunder.m
					// improved by: Kevin van Zonneveld (http://kevin.vanzonneveld.net)
					// improved by: Kevin van Zonneveld (http://kevin.vanzonneveld.net)
					// improved by: Rafał Kukawski (http://kukawski.pl)
					// bugfixed by: Pellentesque Malesuada
					//   example 1: base64_encode('Kevin van Zonneveld');
					//   returns 1: 'S2V2aW4gdmFuIFpvbm5ldmVsZA=='
					//   example 2: base64_encode('a');
					//   returns 2: 'YQ=='

					var b64 = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=';
					var o1, o2, o3, h1, h2, h3, h4, bits, i = 0,
					ac = 0,
					enc = '',
					tmp_arr = [];

					if (!data) {
						return data;
					}

					do { // pack three octets into four hexets
						o1 = data.charCodeAt(i++);
						o2 = data.charCodeAt(i++);
						o3 = data.charCodeAt(i++);

						bits = o1 << 16 | o2 << 8 | o3;

						h1 = bits >> 18 & 0x3f;
						h2 = bits >> 12 & 0x3f;
						h3 = bits >> 6 & 0x3f;
						h4 = bits & 0x3f;

						// use hexets to index into b64, and append result to encoded string
						tmp_arr[ac++] = b64.charAt(h1) + b64.charAt(h2) + b64.charAt(h3) + b64.charAt(h4);
					} while (i < data.length);

					enc = tmp_arr.join('');

					var r = data.length % 3;

					return (r ? enc.slice(0, r - 3) : enc) + '==='.slice(r || 3);
				};

				e.format = function (s, c) {
					return s.replace(/{(\w+)}/g, function (m, p) {
						return c[p];
					});
				};
				e.ctx = {
					worksheet: name || "Worksheet",
					table: table
				};
				window.location.href = e.uri + e.base64_encode(e.format(e.template, e.ctx));
			}
		};

		$.fn[ pluginName ] = function ( options ) {
				this.each(function() {
						if ( !$.data( this, "plugin_" + pluginName ) ) {
								$.data( this, "plugin_" + pluginName, new Plugin( this, options ) );
						}
				});

				// chain jQuery functions
				return this;
		};

})( jQuery, window, document );