# Interview MP3 + Subtitle Generator

Windows에서 면접 질문/답변 대본을 넣으면 아래 결과물을 자동 생성하는 1차 버전입니다.

- 개별 MP3 파일
- 전체 자막 파일 `.srt`
- 가사형 자막 `.lrc`
- 재생목록 `.m3u`
- 파싱 결과 `.csv`

## 1. 설치

```bash
pip install -r requirements.txt
```

## 2. 실행

```bash
python interview_tts_generator.py
```

## 3. 입력 형식

기본적으로 아래 형식을 권장합니다.

```txt
Q: Please introduce yourself.
A: My name is Lee...

Q: Why do you want to join Seapeak?
A: I want to join Seapeak because...
```

지원 파일 형식:
- `.txt`
- `.docx`
- `.csv`

CSV는 아래 헤더 중 하나를 사용하면 됩니다.

- `Q, A`
- `Question, Answer`

## 4. 출력 결과

예시:

- `01_Q01.mp3`
- `02_Q01_Answer.mp3`
- `full_interview_practice.srt`
- `full_interview_practice.lrc`
- `playlist.m3u`
- `parsed_segments.csv`

## 5. 특징

- 면접관 / 지원자 음성 분리 가능
- 속도 / pitch / pause 조절 가능
- 자막은 문장 단위로 자동 분리
- 영어 면접 연습용으로 적합

## 6. 주의점

이 버전은 `edge-tts` 기반입니다.
즉, **실행하는 PC에서 인터넷 연결이 필요할 수 있습니다.**

## 7. 다음 업그레이드 후보

- 하나의 `full_interview_practice.mp3`로 자동 병합
- 질문 후 답변 전 2초 대기 자동 삽입
- Korean meaning 동시 표시용 자막 생성
- Shadowing mode / repeat mode
- GUI에서 직접 문장 편집
- exe 패키징
