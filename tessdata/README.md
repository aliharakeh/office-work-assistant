# OCR language data

`eng.traineddata` is the English LSTM model from
https://github.com/tesseract-ocr/tessdata_fast/tree/87416418657359cb625c412a48b6e1d6d41c29bd

`ara.traineddata` is the Arabic LSTM model from
https://github.com/tesseract-ocr/tessdata_best/tree/e12c65a915945e4c28e237a9b52bc4a8f39a0cec

Both use Apache License 2.0 (see LICENSE.txt). Arabic uses the best model because
the fast model inserted unwanted characters in our 32-bit Arabic smoke test.

Source downloads:
https://raw.githubusercontent.com/tesseract-ocr/tessdata_fast/87416418657359cb625c412a48b6e1d6d41c29bd/eng.traineddata
https://raw.githubusercontent.com/tesseract-ocr/tessdata_best/e12c65a915945e4c28e237a9b52bc4a8f39a0cec/ara.traineddata

SHA-256:
- English: `7d4322bd2a7749724879683fc3912cb542f19906c83bcc1a52132556427170b2`
- Arabic: `ab9d157d8e38ca00e7e39c7d5363a5239e053f5b0dbdb3167dde9d8124335896`

Other language models can be copied into the deployed app's `tessdata` folder.
Use the matching language codes in the OCR step, such as `eng+ara`.
No models are downloaded at runtime.
