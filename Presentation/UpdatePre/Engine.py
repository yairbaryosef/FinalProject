from Presentation.UpdatePre.Slide1 import Slide1Editor
from Presentation.UpdatePre.Slide2 import Slide2Editor

presentationName= "presentation.pptx"
editor = Slide1Editor("Market-Analysis-Yahoo-Investment-Opportunity.pptx")
editor.update_title_and_content("Hey Yair", "Hey hey hey")
editor.save(presentationName)

editor = Slide2Editor(presentationName)

editor.reset_and_update_content({
    "Digital Pioneer": [
        "Founded in 1994, Yahoo evolved into an AI-powered media and tech company.",
        "Acquired by Apollo in 2021 for $5B."
    ],
    "Key Business Segments": [
        "• AI",
        "• Content platforms",
        "• Email services",
        "• Financial media"
    ],
    "Market Position": [
        "Over 920 million monthly users.",
        "Competes in a $500B+ digital advertising market alongside Google and Meta."
    ]
})

editor.save(presentationName)



