# coding: utf-8
import os, re, shutil, sys ,math
import xlrd
from matplotlib import pyplot
import numpy
from datetime import datetime

def excel_date(num):
    from datetime import datetime, timedelta
    return(datetime(1899, 12, 30) + timedelta(days=num))

def serial(s, excel=False):
    import datetime
    end = datetime.datetime.strptime(s, '%Y/%m/%d')
    start = datetime.datetime(1899, 12, 31)
    delta = end - start
    if excel and end >= datetime.datetime(1900, 3, 1):
        return delta.days + 1
    return delta.days

def read_xlsx(filename,start_time,end_time) :
    book = xlrd.open_workbook(filename)
    # ブック内のシート数を取得
    num_of_worksheets = book.nsheets
    # 全シートの名前を取得
    sheet_names = book.sheet_names()

    sheet = book.sheet_by_index(0)
    # キーワード取得
    keys = []
    for j in range(sheet.ncols) :
        val = sheet.cell_value(rowx=0,colx=j)
        keys.append(val)
    keys_keyword = dict(zip(keys,[ [] for i in keys])) # filterのためのキーワードを保管

    all_data = []
    for i in range(1,sheet.nrows) :
        data = {}
        for j,key in enumerate(keys) :
            val = sheet.cell_value(rowx=i,colx=j)
            data[key] = val
            if not val in keys_keyword[key] and not key == u"日付" and not key == u"乱数" :
                keys_keyword[key].append(val)
            if key == u"日付" :
                data[key] = excel_date(val).timestamp()
        all_data.append(data)
    print("all data",len(all_data))    

    all_data = filter(lambda data:data[u"性別"]!=u"", all_data) # 欠損値があるものを削除

    for key in keys_keyword :
        print(key ,len(keys_keyword[key]))

    x = []
    y = []
    counter = 1
    Prefectures = []
    Prefecture = []
    Citys = []
    City = []
    Jobs = []
    Job = []
    Medias = []
    Media = []    
    for i,data in enumerate(sorted(all_data,key=lambda data:data[u"日付"]) ):
        if not i==0 and x[-1] == data[u"日付"] :
            y[-1] += 1
            counter += 1
            if not data[u"都道府県"] in Prefectures[-1] :
                Prefectures[-1].append(data[u"都道府県"])
            if not data[u"市区郡"] in Citys[-1] :
                Citys[-1].append(data[u"市区郡"])
            if not data[u"業種"] in Jobs[-1] :
                Jobs[-1].append(data[u"業種"])
            if not data[u"応募媒体"] in Medias[-1] :
                Medias[-1].append(data[u"応募媒体"])
            continue
        x.append(data[u"日付"])
        y.append(counter)
        Prefectures.append(Prefecture)
        Citys.append(City)
        Jobs.append(Job)
        Medias.append(Media)
        counter = 1 # 1日当たりの求人
        #counter += 1 # 合計
        Prefecture = [] # 応募のあった都道府県
        City = [] # 応募のあった市区郡
        Job = [] # 応募のあった業種
        Media = [] # 応募のあった応募媒体

    # 
    X = []
    Y = []
    if start_time < x[0] :
        old_time = start_time - 86400.0
    else :
        old_time = x[0] - 86400.0        
    for POSIX_time,Prefecture,City,Job,y_value in zip(x,Prefectures,Citys,Jobs,y) :
        while not POSIX_time - old_time == 86400.0 :
            # 応募がなかった日の特徴量
            old_time += 86400.0
            print("no data : ", datetime.fromtimestamp(old_time))
            Y.append(0)
            data = []
            # 時間関係の特徴量
            data.append(old_time)
            time = datetime.fromtimestamp(old_time)
            year = time.year
            data.append(year)
            month = time.month
            data.append(month)
            day = time.day
            data.append(day)
            weekday = time.weekday()
            data.append(weekday)
            # 残りの特徴量
            for s in keys_keyword[u"都道府県"] + keys_keyword[u"市区郡"] + keys_keyword[u"業種"] + keys_keyword[u"応募媒体"] :
                data.append(0)
            X.append( data )
        data = []
        # 時間関係の特徴量
        data.append(POSIX_time)
        time = datetime.fromtimestamp(POSIX_time)
        year = time.year
        data.append(year)
        month = time.month
        data.append(month)
        day = time.day
        data.append(day)
        weekday = time.weekday()
        data.append(weekday)

        # 都道府県の特徴量
        for s in keys_keyword[u"都道府県"] :
            if s in Prefecture :
                data.append(1)
            else :
                data.append(0)

        # 市区郡の特徴量
        for s in keys_keyword[u"市区郡"] :
            if s in City :
                data.append(1)
            else :
                data.append(0)
        
        # 業種の特徴量
        for s in keys_keyword[u"業種"] :
            if s in Job :
                data.append(1)
            else :
                data.append(0)

        # 応募媒体の特徴量
        for s in keys_keyword[u"応募媒体"] :
            if s in Job :
                data.append(1)
            else :
                data.append(0)

        X.append( data )
        Y.append( y_value )
        old_time = POSIX_time

    while old_time < end_time :
        # 応募がなかった日の特徴量
        old_time += 86400.0
        print("no data : ", datetime.fromtimestamp(old_time))
        Y.append(0)
        data = []
        # 時間関係の特徴量
        data.append(old_time)
        time = datetime.fromtimestamp(old_time)
        year = time.year
        data.append(year)
        month = time.month
        data.append(month)
        day = time.day
        data.append(day)
        weekday = time.weekday()
        data.append(weekday)
        # 残りの特徴量
        for s in keys_keyword[u"都道府県"] + keys_keyword[u"市区郡"] + keys_keyword[u"業種"] + keys_keyword[u"応募媒体"] :
            data.append(0)
        X.append( data )

    X = numpy.array(X)
    return X,numpy.array(Y),X.T[0]

def read_xlsx2(filename,start_time,end_time,filter_words) :
    book = xlrd.open_workbook(filename)
    # ブック内のシート数を取得
    num_of_worksheets = book.nsheets
    # 全シートの名前を取得
    sheet_names = book.sheet_names()

    sheet = book.sheet_by_index(0)
    # キーワード取得
    keys = []
    for j in range(sheet.ncols) :
        val = sheet.cell_value(rowx=0,colx=j)
        keys.append(val)
    keys_keyword = dict(zip(keys,[ [] for i in keys])) # filterのためのキーワードを保管

    all_data = []
    for i in range(1,sheet.nrows) :
        data = {}
        for j,key in enumerate(keys) :
            val = sheet.cell_value(rowx=i,colx=j)
            data[key] = val
            if not val in keys_keyword[key] and not key == u"日付" and not key == u"乱数" :
                keys_keyword[key].append(val)
            if key == u"日付" :
                data[key] = excel_date(val).timestamp()
        all_data.append(data)
    print("all data",len(all_data))    

    all_data = list(filter(lambda data:data[u"性別"]!=u"", all_data)) # 欠損値があるものを削除

    for key in keys_keyword :
        print(key ,len(keys_keyword[key]))

    # フィルター
    for key in filter_words :
        if filter_words[key] :
            all_data = list(filter(lambda data:(data[key] in filter_words[key]) ,all_data))

    x = []
    y = []
    counter = 1
    Prefectures = []
    Prefecture = []
    Citys = []
    City = []
    Jobs = []
    Job = []
    Medias = []
    Media = []    
    for i,data in enumerate(sorted(all_data,key=lambda data:data[u"日付"]) ):
        if not i==0 and x[-1] == data[u"日付"] :
            y[-1] += 1
            counter += 1
            if not data[u"都道府県"] in Prefectures[-1] :
                Prefectures[-1].append(data[u"都道府県"])
            if not data[u"市区郡"] in Citys[-1] :
                Citys[-1].append(data[u"市区郡"])
            if not data[u"業種"] in Jobs[-1] :
                Jobs[-1].append(data[u"業種"])
            if not data[u"応募媒体"] in Medias[-1] :
                Medias[-1].append(data[u"応募媒体"])
            continue
        x.append(data[u"日付"])
        y.append(counter)
        Prefectures.append(Prefecture)
        Citys.append(City)
        Jobs.append(Job)
        Medias.append(Media)
        counter = 1 # 1日当たりの求人
        #counter += 1 # 合計
        Prefecture = [] # 応募のあった都道府県
        City = [] # 応募のあった市区郡
        Job = [] # 応募のあった業種
        Media = [] # 応募のあった応募媒体

    # 
    X = []
    Y = []
    if start_time < x[0] :
        old_time = start_time - 86400.0
    else :
        old_time = x[0] - 86400.0        
    for POSIX_time,Prefecture,City,Job,y_value in zip(x,Prefectures,Citys,Jobs,y) :
        while not POSIX_time - old_time == 86400.0 :
            # 応募がなかった日の特徴量
            old_time += 86400.0
            print("no data : ", datetime.fromtimestamp(old_time))
            Y.append(0)
            data = []
            # 時間関係の特徴量
            data.append(old_time)
            time = datetime.fromtimestamp(old_time)
            year = time.year
            data.append(year)
            month = time.month
            data.append(month)
            day = time.day
            data.append(day)
            weekday = time.weekday()
            data.append(weekday)
            X.append( data )
            
        data = []
        # 時間関係の特徴量
        data.append(POSIX_time)
        time = datetime.fromtimestamp(POSIX_time)
        year = time.year
        data.append(year)
        month = time.month
        data.append(month)
        day = time.day
        data.append(day)
        weekday = time.weekday()
        data.append(weekday)
        X.append( data )
        Y.append( y_value )
        old_time = POSIX_time

    while old_time < end_time :
        # 応募がなかった日の特徴量
        old_time += 86400.0
        print("no data : ", datetime.fromtimestamp(old_time))
        Y.append(0)
        data = []
        # 時間関係の特徴量
        data.append(old_time)
        time = datetime.fromtimestamp(old_time)
        year = time.year
        data.append(year)
        month = time.month
        data.append(month)
        day = time.day
        data.append(day)
        weekday = time.weekday()
        data.append(weekday)
        X.append( data )

    X = numpy.array(X)
    return X,numpy.array(Y),X.T[0]

    
def LinReg(X_train,y_train,X_test,y_test):
    # 線形回帰
    # y = sum w_{i}*x_{i}
    from sklearn.linear_model import LinearRegression
    clf = LinearRegression()
    fit = clf.fit(X_train, y_train)
    print("Training score : ",fit.score(X_train, y_train))
    print("Test set score : ",fit.score(X_test, y_test))
    print(fit.coef_)
    y_train_predict = fit.predict(X_train)
    y_test_predict = fit.predict(X_test)
    return fit,y_train_predict,y_test_predict

def ridge(X_train,y_train,X_test,y_test,alpha=10.0):
    # Ridge
    # 線形回帰にL2正則化を追加したもの
    # y = sum w_{i}*x_{i} + |w|^2    
    from sklearn.linear_model import Ridge
    clf = Ridge(alpha=alpha)
    fit = clf.fit(X_train, y_train)
    print("Training score : ",fit.score(X_train, y_train))
    print("Test set score : ",fit.score(X_test, y_test))
    print(fit.coef_)
    print(fit.intercept_) # 切片
    y_train_predict = fit.predict(X_train)
    y_test_predict = fit.predict(X_test)
    return fit,y_train_predict,y_test_predict

def lasso(X_train,y_train,X_test,y_test,alpha=10.0,max_iter=100000):
    # Lasso
    # 線形回帰にL1正則化を追加したもの
    # y = sum w_{i}*x_{i} + |w|
    from sklearn.linear_model import Lasso
    clf = Lasso(alpha=alpha,max_iter=max_iter)
    fit = clf.fit(X_train, y_train)
    print("Training score : ",fit.score(X_train, y_train))
    print("Test set score : ",fit.score(X_test, y_test))
    print(fit.coef_)
    print(fit.intercept_) # 切片
    y_train_predict = fit.predict(X_train)
    y_test_predict = fit.predict(X_test)
    return fit,y_train_predict,y_test_predict

def kRidge(X_train,y_train,X_test,y_test,alpha=10.0) :
    from sklearn.kernel_ridge import KernelRidge
    clf = KernelRidge(alpha=alpha, kernel='rbf')
    fit = clf.fit(X_train, y_train)
    print("Training score : ",fit.score(X_train, y_train))
    print("Test set score : ",fit.score(X_test, y_test))
    y_train_predict = fit.predict(X_train)
    y_test_predict = fit.predict(X_test)
    return fit,y_train_predict,y_test_predict

def DecTreeReg(X_train,y_train,X_test,y_test,max_depth=4) :
    # 決定木
    # 単純な2値問題を繰り返すことで学習を行う。
    # アルゴリズムの性質上、外挿はできない。
    from sklearn.tree import DecisionTreeRegressor
    tree = DecisionTreeRegressor(max_depth=max_depth,random_state=0)
    fit = tree.fit(X_train,y_train)
    print("Training score : ",fit.score(X_train, y_train))
    print("Test set score : ",fit.score(X_test, y_test))
    print(fit.feature_importances_)
    y_train_predict = fit.predict(X_train)
    y_test_predict = fit.predict(X_test)
    return fit,y_train_predict,y_test_predict

def RandForest(X_train,y_train,X_test,y_test):
    # ランダムフォレスト
    # サンプル点をブーストラップサンプリング(復元抽出)を行い、それぞれで決定木を計算し確率予想を平均化する。
    # バカパラレルが可能。n_jobs=-1
    from sklearn.ensemble import RandomForestRegressor
    tree = RandomForestRegressor(n_estimators=10,max_depth=10,random_state=0)
    fit = tree.fit(X_train,y_train)
    print("Training score : ",fit.score(X_train, y_train))
    print("Test set score : ",fit.score(X_test, y_test))
    print(fit.feature_importances_)
    y_train_predict = fit.predict(X_train)
    y_test_predict = fit.predict(X_test)
    return fit,y_train_predict,y_test_predict

def GradBoostReg(X_train,y_train,X_test,y_test) :
    # 勾配ブースティング決定木
    # 決定木を計算し、間違っている部分を新たな決定木で修正する。
    # ランダムフォレストよりも時間がかかる。
    from sklearn.ensemble import GradientBoostingRegressor
    tree = GradientBoostingRegressor(n_estimators=10,max_depth=10,random_state=0)
    fit = tree.fit(X_train,y_train)
    print("Training score : ",fit.score(X_train, y_train))
    print("Test set score : ",fit.score(X_test, y_test))
    print(fit.feature_importances_)
    y_train_predict = fit.predict(X_train)
    y_test_predict = fit.predict(X_test)
    return fit,y_train_predict,y_test_predict

def SVRfunc(X_train,y_train,X_test,y_test):
    # SVRは回帰、SVCは分類
    # 入力データの正規化が必要
    from sklearn.preprocessing import StandardScaler
    scaler = StandardScaler()
    scaler.fit(X_train)
    X_train_scaled = scaler.transform(X_train)
    X_test_scaled = scaler.transform(X_test)
    
    from sklearn.svm import SVR
    tree = SVR(kernel='rbf',C=10,gamma=1.5)
    fit = tree.fit(X_train_scaled,y_train)
    print("Training score : ",fit.score(X_train_scaled, y_train))
    print("Test set score : ",fit.score(X_test_scaled, y_test))
    y_train_predict = fit.predict(X_train_scaled)
    y_test_predict = fit.predict(X_test_scaled)
    return fit,y_train_predict,y_test_predict

def MPLReg(X_train,y_train,X_test,y_test):
    # ニューラルネットワーク
    # 入力データの正規化が必要
    # 入力データの数によってalphaの値を変化させる必要がある。
    from sklearn.preprocessing import StandardScaler
    scaler = StandardScaler()
    scaler.fit(X_train)
    X_train_scaled = scaler.transform(X_train)
    X_test_scaled = scaler.transform(X_test)

    from sklearn.neural_network import MLPRegressor
    tree = MLPRegressor(solver='lbfgs',random_state=0,hidden_layer_sizes=[1000,100,100,100],activation='tanh',alpha=0.1)
    fit = tree.fit(X_train_scaled,y_train)
    print("Training score : ",fit.score(X_train_scaled, y_train))
    print("Test set score : ",fit.score(X_test_scaled, y_test))
    #for i,j in zip(y_test,fit.predict( X_test_scaled ))  :
    #s    print(i,j)
    y_train_predict = fit.predict(X_train_scaled)
    y_test_predict = fit.predict(X_test_scaled)
    return fit,y_train_predict,y_test_predict

def plot(times,y_exact,times_train,y_train_predict,times_test,y_test_predict):
    xtick = []
    locs = []
    for POSIX_time in times :
        time = datetime.fromtimestamp(POSIX_time)
        if time.day == 1 :
            xtick.append("%s"%time)
            locs.append(POSIX_time)
    pyplot.xticks(locs, xtick, color="c", fontsize=8, rotation=-30)
    pyplot.plot(times,y_exact,label="exact")

    pyplot.plot(times_train,y_train_predict,label="train")
    pyplot.plot(times_test,y_test_predict,label="predict")
    pyplot.legend()
    pyplot.show()
    return 0

def averaging(X_train,y_train,X_test,y_test,day=7) :
    # day 分だけ使用した平均値を使用して値を滑らかにする。
    j = int(day/2)
    new_X_train = []
    new_y_train = []
    for i in range(len(y_train)-day) :
        i = i + j
        new_X_train.append( X_train[i] )
        y = numpy.average( y_train[i-j:i+day-j] )
        new_y_train.append( y )
    new_X_test = []
    new_y_test = []
    for i in range(len(y_test)-day) :
        i = i + j
        new_X_test.append( X_test[i] )
        y = numpy.average( y_test[i-j:i+day-j] )
        new_y_test.append( y )
    return numpy.array(new_X_train),numpy.array(new_y_train),numpy.array(new_X_test),numpy.array(new_y_test)
    
if __name__ == "__main__":
    # データがあるのは 2014/10/01 ~ 2015/10/31 。 2015/6/1 当たりで分布が大きく変化してる気がする
    # 応募がない日は0人応募があったとして学習させている
    training_start_time = datetime.strptime("2015/6/1", '%Y/%m/%d').timestamp()
    training_end_time = datetime.strptime("2015/9/30", '%Y/%m/%d').timestamp()
    predic_start_time = datetime.strptime("2015/10/1", '%Y/%m/%d').timestamp()
    predic_end_time = datetime.strptime("2015/10/31", '%Y/%m/%d').timestamp()
    filter_words = {}
    # None だとフィルターなし
    filter_words[u"都道府県"] = [u"東京都"]
    filter_words[u"市区郡"] = None#[u"港区"]
    filter_words[u"業種"] = [u"ファストフード"]
    filter_words[u"応募媒体"] = None #["A1","A2","A3","A4"]

    #X,y_exact,times = read_xlsx("kadai_data.xlsx",training_start_time,predic_end_time)# 都道府県等のデータを特徴量として使用。応募が1以上あった場合は1で、それ以外では0になるようにしているため、予測するときに入力する値が何を意味しているかよくわからない。
    X,y_exact,times = read_xlsx2("kadai_data.xlsx",training_start_time,predic_end_time,filter_words)# 都道府県等のデータをフィルターに使用

    # 時間で区切って学習データとテストデータを分ける
    # 学習データの取得
    y_train = y_exact[ numpy.where( X.T[0] >= training_start_time ) ]
    X_train = X[ numpy.where( X.T[0] >= training_start_time ) ]
    y_train = y_train[ numpy.where( X_train.T[0] <= training_end_time ) ]
    X_train = X_train[ numpy.where( X_train.T[0] <= training_end_time ) ]
    # テストデータの取得
    y_test = y_exact[ numpy.where( X.T[0] >= predic_start_time ) ]
    X_test = X[ numpy.where( X.T[0] >= predic_start_time ) ]
    y_test = y_test[ numpy.where( X_test.T[0] <= predic_end_time ) ]
    X_test = X_test[ numpy.where( X_test.T[0] <= predic_end_time ) ]
    
    #X_train,y_train,X_test,y_test = averaging(X_train,y_train,X_test,y_test,7) # 1週間の平均値を使用して応募数をなだらかにする(データ数は1週間分減る)
    times_train = X_train.T[0]
    times_test = X_test.T[0]

    # POSIX_time を特徴量から削除
    #X_train = numpy.delete(X_train, 0, 1)
    #X_test = numpy.delete(X_test, 0, 1)

    # 学習、予測
    #fit,y_train_predict,y_test_predict = LinReg(X_train,y_train,X_test,y_test)
    #fit,y_train_predict,y_test_predict = ridge(X_train,y_train,X_test,y_test,alpha=10.0)
    #fit,y_train_predict,y_test_predict = lasso(X_train,y_train,X_test,y_test,alpha=10.0,max_iter=100000)
    #fit,y_train_predict,y_test_predict = kRidge(X_train,y_train,X_test,y_test,alpha=1.0)
    #fit,y_train_predict,y_test_predict = DecTreeReg(X_train,y_train,X_test,y_test,max_depth=10)
    #fit,y_train_predict,y_test_predict = RandForest(X_train,y_train,X_test,y_test)
    #fit,y_train_predict,y_test_predict = GradBoostReg(X_train,y_train,X_test,y_test)
    fit,y_train_predict,y_test_predict = SVRfunc(X_train,y_train,X_test,y_test)
    #fit,y_train_predict,y_test_predict = MPLReg(X_train,y_train,X_test,y_test)

    print("実際に来た応募の数 : ",numpy.sum(y_test))
    print("予測した応募の数 : ",numpy.sum(y_test_predict))

    plot(times,y_exact,times_train,y_train_predict,times_test,y_test_predict)
