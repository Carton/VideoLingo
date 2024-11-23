import pandas as pd
import os, sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import re
from rich.panel import Panel
from rich.console import Console
import autocorrect_py as autocorrect

console = Console()

CLEANED_CHUNKS_FILE = 'output/log/cleaned_chunks.xlsx'
TRANSLATION_RESULTS_FOR_SUBTITLES_FILE = 'output/log/translation_results_for_subtitles.xlsx'
TRANSLATION_RESULTS_REMERGED_FILE = 'output/log/translation_results_remerged.xlsx'

OUTPUT_DIR = 'output'
AUDIO_OUTPUT_DIR = 'output/audio'

SRC_ONLY_SUBTITLE_OUTPUT_CONFIGS = [
    ('src.srt', ['Source'])
]

TRANSLATION_SUBTITLE_OUTPUT_CONFIGS = [ 
    ('src.srt', ['Source']),
    ('trans.srt', ['Translation']),
    ('src_trans.srt', ['Source', 'Translation']),
    ('trans_src.srt', ['Translation', 'Source'])
]

SRC_ONLY_AUDIO_SUBTITLE_OUTPUT_CONFIGS = [
    ('src_subs_for_audio.srt', ['Source'])
]

TRANSLATION_AUDIO_SUBTITLE_OUTPUT_CONFIGS = [
    ('src_subs_for_audio.srt', ['Source']),
    ('trans_subs_for_audio.srt', ['Translation'])
]

def convert_to_srt_format(start_time, end_time):
    """Convert time (in seconds) to the format: hours:minutes:seconds,milliseconds"""
    def seconds_to_hmsm(seconds):
        hours = int(seconds // 3600)
        minutes = int((seconds % 3600) // 60)
        seconds = seconds % 60
        milliseconds = int(seconds * 1000) % 1000
        return f"{hours:02d}:{minutes:02d}:{int(seconds):02d},{milliseconds:03d}"

    start_srt = seconds_to_hmsm(start_time)
    end_srt = seconds_to_hmsm(end_time)
    return f"{start_srt} --> {end_srt}"

def remove_punctuation(text):
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'[^\w\s]', '', text)
    return text.strip()

def show_difference(str1, str2):
    """Show the difference positions between two strings"""
    min_len = min(len(str1), len(str2))
    diff_positions = []
    
    for i in range(min_len):
        if str1[i] != str2[i]:
            diff_positions.append(i)
    
    if len(str1) != len(str2):
        diff_positions.extend(range(min_len, max(len(str1), len(str2))))
    
    print("Difference positions:")
    print(f"Expected sentence: {str1}")
    print(f"Actual match: {str2}")
    print("Position markers: " + "".join("^" if i in diff_positions else " " for i in range(max(len(str1), len(str2)))))
    print(f"Difference indices: {diff_positions}")

def get_sentence_timestamps(df_words, df_sentences):
    time_stamp_list = []
    
    # Build complete string and position mapping
    full_words_str = ''
    position_to_word_idx = {}
    
    for idx, word in enumerate(df_words['text']):
        clean_word = remove_punctuation(word.lower())
        start_pos = len(full_words_str)
        full_words_str += clean_word
        for pos in range(start_pos, len(full_words_str)):
            position_to_word_idx[pos] = idx
    
    current_pos = 0
    for idx, sentence in df_sentences['Source'].items():
        clean_sentence = remove_punctuation(sentence.lower()).replace(" ", "")
        sentence_len = len(clean_sentence)
        
        match_found = False
        while current_pos <= len(full_words_str) - sentence_len:
            if full_words_str[current_pos:current_pos+sentence_len] == clean_sentence:
                start_word_idx = position_to_word_idx[current_pos]
                end_word_idx = position_to_word_idx[current_pos + sentence_len - 1]
                
                time_stamp_list.append((
                    float(df_words['start'][start_word_idx]),
                    float(df_words['end'][end_word_idx])
                ))
                
                current_pos += sentence_len
                match_found = True
                break
            current_pos += 1
            
        if not match_found:
            print(f"\n⚠️ Warning: No exact match found for sentence: {sentence}")
            show_difference(clean_sentence, 
                          full_words_str[current_pos:current_pos+len(clean_sentence)])
            print("\nOriginal sentence:", df_sentences['Source'][idx])
            raise ValueError("❎ No match found for sentence.")
    
    return time_stamp_list

def align_timestamp(df_text, df_translate, subtitle_output_configs: list, output_dir: str, for_display: bool = True):
    """Align timestamps and add a new timestamp column to df_translate"""
    df_trans_time = df_translate.copy()

    # Assign an ID to each word in df_text['text'] and create a new DataFrame
    words = df_text['text'].str.split(expand=True).stack().reset_index(level=1, drop=True).reset_index()
    words.columns = ['id', 'word']
    words['id'] = words['id'].astype(int)

    # Process timestamps ⏰
    time_stamp_list = get_sentence_timestamps(df_text, df_translate)
    df_trans_time['timestamp'] = time_stamp_list
    df_trans_time['duration'] = df_trans_time['timestamp'].apply(lambda x: x[1] - x[0])

    # Remove gaps 🕳️
    for i in range(len(df_trans_time)-1):
        delta_time = df_trans_time.loc[i+1, 'timestamp'][0] - df_trans_time.loc[i, 'timestamp'][1]
        if 0 < delta_time < 1:
            df_trans_time.at[i, 'timestamp'] = (df_trans_time.loc[i, 'timestamp'][0], df_trans_time.loc[i+1, 'timestamp'][0])

    # Convert start and end timestamps to SRT format
    df_trans_time['timestamp'] = df_trans_time['timestamp'].apply(lambda x: convert_to_srt_format(x[0], x[1]))

    # Polish subtitles: replace punctuation in Translation if for_display and column exists
    if for_display and 'Translation' in df_trans_time.columns:
        df_trans_time['Translation'] = df_trans_time['Translation'].apply(lambda x: re.sub(r'[，。]', ' ', x).strip())

    # Output subtitles 📜
    def generate_subtitle_string(df, columns):
        subtitle_lines = []
        for i, row in df.iterrows():
            subtitle_lines.append(f"{i+1}")
            subtitle_lines.append(row['timestamp'])

            # 添加每个指定列的内容，如果该列存在的话
            for col in columns:
                if col in row:
                    subtitle_lines.append(row[col].strip())

            subtitle_lines.append("")  # 空行分隔

        return '\n'.join(subtitle_lines).strip()

    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
        for filename, columns in subtitle_output_configs:
            # 检查所需的列是否都存在
            if all(col in df_trans_time.columns for col in columns):
                subtitle_str = generate_subtitle_string(df_trans_time, columns)
                with open(os.path.join(output_dir, filename), 'w', encoding='utf-8') as f:
                    f.write(subtitle_str)
            else:
                console.print(f"[yellow]Warning: Skipping {filename} as some required columns are missing[/yellow]")

    return df_trans_time

# ✨ Beautify the translation
def clean_translation(x):
    if pd.isna(x):
        return ''
    cleaned = str(x).strip('。').strip('，')
    return autocorrect.format(cleaned)

def align_timestamp_main(need_translation=True):
    df_text = pd.read_excel(CLEANED_CHUNKS_FILE)
    df_text['text'] = df_text['text'].str.strip('"').str.strip()
    
    if need_translation:
        # 原有的翻译处理逻辑
        df_translate = pd.read_excel(TRANSLATION_RESULTS_FOR_SUBTITLES_FILE)
        df_translate['Translation'] = df_translate['Translation'].apply(clean_translation)
        
        align_timestamp(df_text, df_translate, TRANSLATION_SUBTITLE_OUTPUT_CONFIGS, OUTPUT_DIR)
        console.print(Panel("[bold green]🎉📝 Subtitles generation completed! Please check in the `output` folder 👀[/bold green]"))

        # for audio
        df_translate_for_audio = pd.read_excel(TRANSLATION_RESULTS_REMERGED_FILE)
        df_translate_for_audio['Translation'] = df_translate_for_audio['Translation'].apply(clean_translation)
        
        align_timestamp(df_text, df_translate_for_audio, TRANSLATION_AUDIO_SUBTITLE_OUTPUT_CONFIGS, AUDIO_OUTPUT_DIR)
        console.print(Panel("[bold green]🎉📝 Audio subtitles generation completed! Please check in the `output/audio` folder 👀[/bold green]"))
    else:
        # 只处理原语言字幕
        df_source = pd.read_excel(TRANSLATION_RESULTS_FOR_SUBTITLES_FILE)
        
        align_timestamp(df_text, df_source, SRC_ONLY_SUBTITLE_OUTPUT_CONFIGS, OUTPUT_DIR)
        console.print(Panel("[bold green]🎉📝 Source subtitles generation completed! Please check in the `output` folder 👀[/bold green]"))

        # for audio
        df_source_for_audio = pd.read_excel(TRANSLATION_RESULTS_REMERGED_FILE)
        
        align_timestamp(df_text, df_source_for_audio, SRC_ONLY_AUDIO_SUBTITLE_OUTPUT_CONFIGS, AUDIO_OUTPUT_DIR)
        console.print(Panel("[bold green]🎉📝 Source audio subtitles generation completed! Please check in the `output/audio` folder 👀[/bold green]"))
    

if __name__ == '__main__':
    align_timestamp_main()