/* global $, Word */
/* jshint esversion: 6 */

let releaseContent = {};

function getReleaseDisplay(releaseContent) {
    return 'Release <b>' + releaseContent.Version + '</b>';
}

function storeAndDisplayRelease(rct) {
    releaseContent = JSON.parse(rct);
    $('.release-info').html(getReleaseDisplay(releaseContent));
}

async function insertRelease() {
    console.log(releaseContent);
    Word.run(async function (context) {
	var body = context.document.body;

	context.load(body, 'text');

	await context.sync();

	console.log(body.text);
    });
}

async function init(reason) {
    $(document).ready(function () {
	$('#release_json').on('change', function () {
	    let file = $(this)[0].files[0];
	    let fr = new FileReader();
	    fr.onload = function () {
		storeAndDisplayRelease(fr.result);
	    };
	    fr.readAsText(file);
	});
	$('.insert-release').click(function (e) {
	    e.preventDefault();

	    insertRelease();
	    
	    return false;
	});
    });
}


export default {
    init: init
};
