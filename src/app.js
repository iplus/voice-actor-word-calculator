import docx4js from "docx4js"
import Dropzone from "dropzone";
import "dropzone/dist/dropzone.css";

const myDropzone = new Dropzone("#my-form", {
    acceptedFiles: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    autoProcessQueue: false,
    autoQueue: false,
    disablePreviews: true,
});

function extractText(node) {
    if (node.type === 'text') {
        return node.data
    }
    return node.children.map(extractText).join('');
}

function clearText(text) {
    return text.replace(/\(.+?\)|\[.+?\]|\*.+?\*/g, '')
}

function calcWords(text) {
    const result = text.split(/[\s,.\n]/).filter(v => v.replace(/[^а-яА-Яa-zA-Z]/g, '').trim() !== '')
    return result.length
}

myDropzone.on("addedfile", (file) => {
    const result = document.getElementById('result')
    const debug = document.getElementById('debug')
    result.innerHTML = ''
    debug.innerHTML = ''
    myDropzone.removeAllFiles()
    try {
        docx4js.load(file).then(docx => {
            const words = {}
            const parser = {
                createElement() {
                },
                emit() {
                },
                ontr(model, ...other) {
                    const actor = extractText(model.children[1]).trim()
                    const text = extractText(model.children[2])
                    words[actor] = words[actor] || ''
                    words[actor] += ' ' + clearText(text)
                    const tr = document.createElement('tr')
                    const actorName = document.createElement('td')
                    actorName.innerHTML = actor
                    tr.appendChild(actorName)
                    const actorCount = document.createElement('td')
                    actorCount.innerHTML = calcWords(clearText(text))
                    tr.appendChild(actorCount)
                    const actorText = document.createElement('td')
                    actorText.innerHTML = text
                    tr.appendChild(actorText)
                    debug.appendChild(tr)
                }
            }
            docx.parse(parser)
            for (const actor in words) {
                const actorName = document.createElement('b')
                actorName.innerHTML = actor + ': '
                result.appendChild(actorName)
                const actorValue = document.createElement('span')
                actorValue.innerHTML = calcWords(words[actor])
                result.appendChild(actorValue)
                const br = document.createElement('br')
                result.appendChild(br)
            }
        }).catch(e => {
            result.innerHTML = e.stack
        })
    } catch (e) {
        result.innerHTML = e.stack
    }
})
