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

async function insertReleaseHeading(current, context, version) {
    let headingParagraph = current.insertParagraph('Release v ' + version, Word.InsertLocation.after);
    await context.sync();
    headingParagraph.insertBreak('Page', Word.InsertLocation.before);
    headingParagraph.styleBuiltIn = 'Heading1';
    await context.sync();
    return headingParagraph;
}

const months = [
    'January',
    'February',
    'March',
    'April',
    'May',
    'June',
    'July',
    'August',
    'September',
    'October',
    'November',
    'December'
];

async function insertReleaseDate(current, context) {
    const currentDate = new Date();
    
    let releaseDateParagraph = current.insertParagraph(
	'Release date: ' +
	    months[currentDate.getMonth()] +
	    ' ' + currentDate.getDate() + ', ' + currentDate.getFullYear()
	, Word.InsertLocation.after);
    await context.sync();
    releaseDateParagraph.styleBuiltIn = 'Normal';
    await context.sync();
    return releaseDateParagraph;
}

async function insertWorkItemGroup(current, context, headingText, itemsGrouped, itemType) {
    if (itemsGrouped.length <= 0) {
	// don't even bother with the section
	return current;
    }
    let headingParagraph = current.insertParagraph(headingText, Word.InsertLocation.after);
    await context.sync();
    headingParagraph.styleBuiltIn = 'Heading2';
    await context.sync();
    current = headingParagraph;
    for (let subGroup of itemsGrouped) {
	let innerHeading = current.insertParagraph(subGroup.Heading, Word.InsertLocation.after);
	innerHeading.styleBuiltIn = 'Heading3';
	current = innerHeading;
	await context.sync();
	for (let wi of subGroup.Items) {
	    let itemParagraph = current.insertParagraph(itemType + ' #' + wi.ID + ': ' + wi.Title, Word.InsertLocation.after);
	    itemParagraph.styleBuiltIn = 'Normal';
	    current = itemParagraph;
	    await context.sync();
	}
	await context.sync();
    }
    return current;
}

async function insertGeneral(current, context, requiredLayers, previousLayers) {
    let generalParagraph = current.insertParagraph('General', Word.InsertLocation.after);
    generalParagraph.styleBuiltIn = 'Heading3';
    await context.sync();
    let followingLayers = generalParagraph.insertParagraph('This release requires the following layers to be updated:', Word.InsertLocation.after);
    followingLayers.styleBuiltIn = 'Normal';
    await context.sync();
    current = followingLayers;
    current = current.insertParagraph('If you are deploying on an environment not yet running version ' + previousLayers.PreviousVersion + ', please use the following packages for other layers (if applicable):', Word.InsertLocation.after);
    current.styleBuiltIn = 'Normal';
    await context.sync();
    let ranges = current.getTextRanges([','], false);
    await context.sync();
    let firstRange = ranges.getFirst();
    firstRange.load('font');
    await context.sync();
    firstRange.font.underline = 'Single';
    await context.sync();
    
    return current;
}

async function insertInstallNotes(current, context, releaseContent) {
    current = current.insertParagraph('Install Notes', Word.InsertLocation.after);
    current.styleBuiltIn = 'Heading2';
    await context.sync();
    current = await insertGeneral(current, context, releaseContent.RequiredLayers, releaseContent.LayersFromVersion);
    return current
}

function getWorkItemsUnderHeadings(releaseContent, type) {
    let filteredItems = releaseContent.SourceWorkItems.filter(wi => wi.Type === type);

    let mappingDict = {};
    releaseContent.AreaPathMappings.forEach(apm => {
	mappingDict[apm.AreaPath] = apm.Title;
    });
    let getAreaPathHeading = wi => {
	if (mappingDict[wi.AreaPath]) {
	    return mappingDict[wi.AreaPath];
	} else {
	    let split = wi.AreaPath.split('\\');
	    return split[split.length - 1];
	}
    };

    let withHeading = filteredItems.map(wi => {
	return {
	    ID: wi.ID,
	    Title: wi.Title,
	    Heading: getAreaPathHeading(wi)
	};
    });

    let titlesSeen = {};
    let headedList = [];
    withHeading.forEach(function (wi) {
	if (!titlesSeen[wi.Heading]) {
	    titlesSeen[wi.Heading] = [];
	    headedList.push({
		Heading: wi.Heading,
		Items: titlesSeen[wi.Heading]
	    });
	}
	titlesSeen[wi.Heading].push(wi);
    });
    
    return headedList;
}

async function insertReleaseAfterHeading(previousReleaseHeading, releaseContent, context) {
    // previous heading OOXML
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
    
    let headingParagraph = await insertReleaseHeading(current, context, releaseContent.Version);
    let releaseDate = await insertReleaseDate(headingParagraph, context);

    let changeRequests = getWorkItemsUnderHeadings(releaseContent, 'Change Request');
    console.log(changeRequests);
    let bugs = getWorkItemsUnderHeadings(releaseContent, 'Bug');
    console.log(bugs);

    let wiCurrent = await insertWorkItemGroup(releaseDate, context, 'New Features & Improvements', changeRequests, 'Change Request');

    wiCurrent = await insertWorkItemGroup(wiCurrent, context, 'Bug Fixing', bugs, 'Bug');

    let installNotes = await insertInstallNotes(wiCurrent, context, releaseContent)
    
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
