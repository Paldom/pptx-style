# pptx-style README demo. Render from the repo root:  vhs demo/demo.tape
# Shows the real scripts the two skills run, on generated brand-free fixtures;
# each step starts with the plain-English ask that triggers the skill.
Output demo/demo.gif

Require python3

Set Shell "bash"
Set FontSize 16
Set Width 1280
Set Height 720
Set Padding 24
Set Theme "GitHub Dark"
Set TypingSpeed 28ms
Set Framerate 12
Set WaitTimeout 60s

Hide
Type "source demo/setup.sh && clear"
Enter
Sleep 2s
Show

Type "# You: 'profile messy.potx - what is in it, and is it AI-ready?'  (pptx-style-profile)"
Enter
Type "python3 $PROFILE/profile.py messy.potx"
Enter
Wait+Screen /wrote .pptx-style/
Sleep 1500ms
Type "grep -A9 '^## Patterns' .pptx-style/messy/STYLE.md | cut -c1-118"
Enter
Sleep 4s

Type "clear"
Enter
Type "# You: 'build the 18-slide ops review in that template'  (pptx-style-build)"
Enter
Type "python3 $BUILD/build.py spec.json"
Enter
Wait+Screen /lint: \d+ errors/
Sleep 4s

Type "clear"
Enter
Type "# You: 'make the title neon pink in Comic Sans, disclaimer box over the footer'"
Enter
Type "python3 $BUILD/build.py offcontract.json"
Enter
Wait+Screen /nothing written/
Sleep 4s

Type "clear"
Enter
Type "# You: 'design shipped template v2 - did anything break?'  (pptx-style-profile)"
Enter
Type "python3 $PROFILE/diff.py .pptx-style/messy/contract.json drift.potx"
Enter
Wait+Screen /ADDITIVE/
Sleep 5s
