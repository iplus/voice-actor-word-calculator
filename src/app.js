import docx4js from "docx4js"
import Dropzone from "dropzone";
import "dropzone/dist/dropzone.css";

const myDropzone = new Dropzone("#my-form", {
    acceptedFiles: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    autoProcessQueue: false,
    autoQueue: false,
    disablePreviews: true,
    uploadMultiple: false,
    dictDefaultMessage: 'Drop "DOCX" file here',
    parallelUploads: 1
});

function extractText(node) {
    if (!node || !node.type) {
        return ''
    }
    if (node.type === 'text') {
        return node.data
    }
    return node.children.map(extractText).join('');
}

function clearText(text) {
    return text.replace(/\(.+?\)|\[.+?\]|\*.+?\*/g, '')
}

function calcWords(text) {
    const result = text.split(/[^-\w\da-яА-ЯёЁ]/).filter(v => v.replace(/[^\dёЁа-яА-Яa-zA-Z]/g, '').trim() !== '')
    return result.length
}

let loading = false
myDropzone.on("addedfile", (file) => {
    if (loading) return
    loading = true
    const result = document.getElementById('result')
    const debug = document.getElementById('debug')
    result.innerHTML = `<h1>${file.name}</h1><br><br>`
    debug.innerHTML = ''
    try {
        docx4js.load(file).then(docx => {
            const words = {}
            const parser = {
                createElement() {
                },
                emit() {
                },
                ontr(model) {
                    if (model.children.length !== 3 || !extractText(model.children[0]).match(/\d+:\d+/)) {
                        return
                    }
                    const time = extractText(model.children[0]).trim()
                    const actor = extractText(model.children[1]).trim()
                    const text = extractText(model.children[2])
                    words[actor] = words[actor] || ''
                    words[actor] += ' ' + clearText(text)
                    const tr = document.createElement('tr')
                    const actorTime = document.createElement('td')
                    actorTime.innerHTML = time
                    tr.appendChild(actorTime)
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
            if (!Object.keys(words).length) {
                result.innerHTML = 'No tables found in document'
            }
        }).catch(e => {
            result.innerHTML = e.stack
        }).finally(() => {
            loading = false
            myDropzone.removeAllFiles(true)
        })
    } catch (e) {
        result.innerHTML = e.stack
        loading = false
        myDropzone.removeAllFiles(true)
    }
})
