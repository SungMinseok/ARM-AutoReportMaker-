import os
import glob
#file_list = []
directory = fr'D:\Builds\KR'

def get_latest_matching_file(directory, prefix, extension):
    search_pattern = os.path.join(directory, f"{prefix}*.{extension}")
    files = glob.glob(search_pattern)
    
    if not files:
        return None

    # 파일을 수정 시간을 기준으로 정렬하여 최신 파일 가져오기 getctime
    # 만든 시간 : getmtime
    latest_file = max(files, key=os.path.getmtime)
    #files.append(latest_file)
    return os.path.basename(latest_file)
    file_list.append(os.path.basename(latest_file))
    #return latest_file

#prefix = 'R2MClientKorea_Alpha'

# .apk 확장자인 파일 찾기
def get_filename(nation):
    #nation = Korea
    file_list = []
    file_list.append(get_latest_matching_file(directory, f'R2MClient{nation}_Alpha', 'apk'))
    file_list.append(get_latest_matching_file(directory, f'R2MClient{nation}_Alpha', 'ipa'))
    file_list.append(get_latest_matching_file(directory, f'R2MClient{nation}_Live_QA', 'apk'))
    file_list.append(get_latest_matching_file(directory, f'signed_gxs_R2MClient{nation}_Live', 'apk'))
    file_list.append(get_latest_matching_file(directory, f'R2MClient{nation}_Live_QA', 'ipa'))
    #file_list.append(get_latest_matching_file(directory, f'R2MClient{nation}_Alpha', 'ipa'))
    file_list.append('TestFlight 1.x.x(y)')
    return file_list


if __name__ == "__main__" :
    get_latest_matching_file(directory, f'R2MClientKorea_Alpha', 'apk')
    get_latest_matching_file(directory, f'R2MClientKorea_Alpha', 'ipa')
    get_latest_matching_file(directory, f'R2MClientKorea_Live_QA', 'apk')
    get_latest_matching_file(directory, f'signed_gxs_R2MClientKorea_Live', 'apk')
    get_latest_matching_file(directory, f'R2MClientKorea_Live_QA', 'ipa')
    get_latest_matching_file(directory, f'R2MClientKorea_Alpha', 'ipa')


    print(file_list)