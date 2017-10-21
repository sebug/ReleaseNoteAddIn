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

function getVersionFromHeadingParagraph(h) {
    let rvRe = /(\d+\.\d+\.\d+\.\d+)/;
    return rvRe.exec(h.text)[1];
}

function compareVersions(a, b) {
    var aParts = a.split('.');
    var bParts = b.split('.');
    if (aParts.length != bParts.length) {
	throw new Error("Expected both version numbers to be the same size");
    }
    for (let i = 0; i < aParts.length; i += 1) {
	let ap = Number(aParts[i]);
	let bp = Number(bParts[i]);
	if (ap < bp) {
	    return -1;
	} else if (ap > bp) {
	    return 1;
	}
    }
    return 0;
}

async function insertReleaseAfterHeading(previousReleaseHeading, releaseContent, context) {
    // previous heading OOXML
    let previousHeadingOOXML = previousReleaseHeading.getOoxml();
    await context.sync();
    console.log('inserting release after heading...');
    let current = previousReleaseHeading;
    let next = current.getNextOrNullObject();
    if (next) {
	next.load('styleBuiltIn');
	await context.sync();
    }
    while (next && next.styleBuiltIn !== 'Heading1') {
	current = next;
	next = current.getNextOrNullObject();
	if (next) {
	    next.load('styleBuiltIn');
	    await context.sync();
	    let html = next.getHtml();
	    await context.sync();
	    if (!html.value) {
		next = null;
	    } else if (html.value.indexOf('Known Issues and Future Improvements') >= 0) {
		// that's another way to stop
		next = null;
	    }
	}
    }
    console.log('Just before');
    let headingParagraph = current.insertParagraph('Release v ' + releaseContent.Version, Word.InsertLocation.after);
    await context.sync();
    console.log(headingParagraph);
    return true;
}

function findLastLowerReleaseHeading(releaseHeadings, version) {
    let previous;
    // find first heading with a bigger release number, then return the previous
    for (let h of releaseHeadings) {
	let v = getVersionFromHeadingParagraph(h);
	if (v && compareVersions(v, version) > 0) {
	    return previous;
	}
	previous = h;
    }

    // If nothing is found, just return the last one
    return previous;
}

async function insertRelease() {
    console.log(releaseContent);
    Word.run(async function (context) {
	try {
	    let document = context.document;

	    context.load(document, 'sections');

	    await context.sync();

	    let sections = document.sections;

	    sections.load('items');

	    await context.sync();

	    const releaseHeadings = [];

	    for (let section of sections.items) {
		let body = section.body;
		body.load('paragraphs');
		await context.sync();

		for (let p of body.paragraphs.items) {
		    if (p.styleBuiltIn === 'Heading1') {
			releaseHeadings.push(p);
		    }
		}
	    }

	    let headedReleases = releaseHeadings.map(getVersionFromHeadingParagraph);

	    if (headedReleases.filter(r => r === releaseContent.Version).length) {
		$('.insert-message').html('Release has already been inserted.');
	    } else {
		$('.insert-message').html('Inserting release...');
		let previousHeading = findLastLowerReleaseHeading(releaseHeadings, releaseContent.Version);
		if (!previousHeading) {
		    // TODO, that means we don't have any releases in the file yet
		} else {
		    await insertReleaseAfterHeading(previousHeading, releaseContent, context);
		}
	    }
	} catch (e) {
	    console.log(e);
	}
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
