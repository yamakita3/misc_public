#v3ではSHIFTとFLANNに距離によるフィルタリングとしして、距離によるフィルタリングを導入した。
# usage: python3 export_tie_points_v3.py test_mai_org_wgs_4326.png test_mai_org_wgs_4326.pgw test_mai_part2.png test_mai_part2.png.points,0.8
import cv2
import numpy as np
import sys

def process_images(geoimg_file, pgw_file, non_geoimg_file, output_file, distance_threshold):
    # 画像の読み込み
    geoimg_img = cv2.imread(geoimg_file)
    non_geoimg_img = cv2.imread(non_geoimg_file)

    # 特徴点検出
    sift = cv2.SIFT_create()
    keypoints_geoimg, descriptors_geoimg = sift.detectAndCompute(geoimg_img, None)
    keypoints_non_geoimg, descriptors_non_geoimg = sift.detectAndCompute(non_geoimg_img, None)

    # FLANNベースのマッチング
    # FLANN_INDEX_KDTREE = 1
    # index_params = dict(algorithm=FLANN_INDEX_KDTREE, trees=5)
    # search_params = dict(checks=50)
    # flann = cv2.FlannBasedMatcher(index_params, search_params)
    # matches = flann.knnMatch(descriptors_geoimg, descriptors_non_geoimg, k=2)

    # Brute-Force Matcherに変更
    bf = cv2.BFMatcher(cv2.NORM_L2)
    matches = bf.knnMatch(descriptors_geoimg, descriptors_non_geoimg, k=2)

    # 距離によるフィルタリング
    good_matches = []
    for m, n in matches:
        if m.distance < distance_threshold * n.distance:
            good_matches.append(m)

    # マッチングした特徴点を取得
    points_geoimg = np.float32([keypoints_geoimg[m.queryIdx].pt for m in good_matches]).reshape(-1, 1, 2)
    points_non_geoimg = np.float32([keypoints_non_geoimg[m.trainIdx].pt for m in good_matches]).reshape(-1, 1, 2)

    # ワールドファイルから座標を取得
    with open(pgw_file, "r") as f:
        lines = f.readlines()
        pixel_size_x = float(lines[0])
        rotation_x = float(lines[1])
        rotation_y = float(lines[2])
        pixel_size_y = float(lines[3])
        origin_x = float(lines[4])
        origin_y = float(lines[5])

    # 変換行列の計算
    M = np.array([[pixel_size_x, rotation_x, origin_x],
                  [rotation_y, pixel_size_y, origin_y],
                  [0, 0, 1]])

    # 出力ファイルに書き込み
    with open(output_file, 'w') as f:
        f.write("mapX,mapY,sourceX,sourceY,enable\n")
        for i in range(len(points_geoimg)):
            x, y = points_geoimg[i][0]
            source_x, source_y = points_non_geoimg[i][0]
            enable = 0
            # 変換行列を用いて座標変換
            transformed_coordinates = np.dot(M, np.array([x, y, 1]))
            f.write(f"{transformed_coordinates[0]},{transformed_coordinates[1]},{source_x},-{source_y},{enable}\n")

if __name__ == "__main__":
    if len(sys.argv) != 6:
        print("Usage: python3 export_tie_points_v3.py geoimg_file pgw_file non_geoimg_file output_file distance_threshold")
    else:
        geoimg_file = sys.argv[1]
        pgw_file = sys.argv[2]
        non_geoimg_file = sys.argv[3]
        output_file = sys.argv[4]
        distance_threshold = float(sys.argv[5])
        process_images(geoimg_file, pgw_file, non_geoimg_file, output_file, distance_threshold)
# 終了

