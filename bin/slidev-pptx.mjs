#!/usr/bin/env node
import { main } from '../pptx/cli.mjs'

main(process.argv.slice(2)).then(code => { process.exitCode = code }, error => {
  console.error(error.message)
  process.exitCode = 1
})
