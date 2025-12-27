/*
 * 文件名: modelparameter.cpp
 * 文件作用: 项目参数单例类实现文件
 * 功能描述:
 * 1. 实现了分离式保存逻辑：基础参数存 .pwt，图表数据存 _chart.json。
 * 2. 实现了加载时的自动关联读取，确保数据能被恢复。
 */

#include "modelparameter.h"
#include <QFile>
#include <QJsonDocument>
#include <QFileInfo>
#include <QDebug>

ModelParameter* ModelParameter::m_instance = nullptr;

ModelParameter::ModelParameter(QObject* parent) : QObject(parent), m_hasLoaded(false)
{
    m_phi = 0.05; m_h = 20.0; m_mu = 0.5; m_B = 1.05; m_Ct = 5e-4; m_q = 50.0; m_rw = 0.1;
}

ModelParameter* ModelParameter::instance()
{
    if (!m_instance) m_instance = new ModelParameter();
    return m_instance;
}

void ModelParameter::setParameters(double phi, double h, double mu, double B, double Ct, double q, double rw, const QString& path)
{
    m_phi = phi; m_h = h; m_mu = mu; m_B = B; m_Ct = Ct; m_q = q; m_rw = rw;
    m_projectFilePath = path;

    QFileInfo fi(path);
    m_projectPath = fi.isFile() ? fi.absolutePath() : path;
    m_hasLoaded = true;

    if (m_fullProjectData.isEmpty()) {
        QJsonObject reservoir;
        reservoir["porosity"] = m_phi;
        reservoir["thickness"] = m_h;
        reservoir["wellRadius"] = m_rw;
        reservoir["productionRate"] = m_q;
        QJsonObject pvt;
        pvt["viscosity"] = m_mu;
        pvt["volumeFactor"] = m_B;
        pvt["compressibility"] = m_Ct;
        m_fullProjectData["reservoir"] = reservoir;
        m_fullProjectData["pvt"] = pvt;
    }
}

// 辅助函数：构造数据文件路径
// 规则：原文件名 + "_chart.json"
QString ModelParameter::getPlottingDataFilePath() const
{
    if (m_projectFilePath.isEmpty()) return QString();
    QFileInfo fi(m_projectFilePath);
    QString baseName = fi.completeBaseName(); // 不带后缀的文件名
    return fi.absolutePath() + "/" + baseName + "_chart.json";
}

bool ModelParameter::loadProject(const QString& filePath)
{
    // 1. 加载主项目文件 (.pwt)
    QFile file(filePath);
    if (!file.open(QIODevice::ReadOnly)) return false;

    QByteArray data = file.readAll();
    file.close();

    QJsonDocument doc = QJsonDocument::fromJson(data);
    if (doc.isNull()) return false;

    m_fullProjectData = doc.object();

    // 解析物理参数
    if (m_fullProjectData.contains("reservoir")) {
        QJsonObject res = m_fullProjectData["reservoir"].toObject();
        m_q = res["productionRate"].toDouble(50.0);
        m_phi = res["porosity"].toDouble(0.05);
        m_h = res["thickness"].toDouble(20.0);
        m_rw = res["wellRadius"].toDouble(0.1);
    }
    if (m_fullProjectData.contains("pvt")) {
        QJsonObject pvt = m_fullProjectData["pvt"].toObject();
        m_Ct = pvt["compressibility"].toDouble(5e-4);
        m_mu = pvt["viscosity"].toDouble(0.5);
        m_B = pvt["volumeFactor"].toDouble(1.05);
    }

    m_projectFilePath = filePath;
    m_projectPath = QFileInfo(filePath).absolutePath();
    m_hasLoaded = true;

    // 2. [新增] 尝试加载同名的图表数据文件 (_chart.json)
    QString dataFilePath = getPlottingDataFilePath();
    QFile dataFile(dataFilePath);
    if (dataFile.exists() && dataFile.open(QIODevice::ReadOnly)) {
        QJsonDocument dataDoc = QJsonDocument::fromJson(dataFile.readAll());
        if (!dataDoc.isNull() && dataDoc.isObject()) {
            QJsonObject dataObj = dataDoc.object();
            // 将读取到的 plotting_data 放入内存缓存中，供 getPlottingData 调用
            if (dataObj.contains("plotting_data")) {
                m_fullProjectData["plotting_data"] = dataObj["plotting_data"];
                qDebug() << "成功加载图表数据文件:" << dataFilePath;
            }
        }
        dataFile.close();
    } else {
        qDebug() << "未找到图表数据文件(可能是新项目):" << dataFilePath;
    }

    return true;
}

bool ModelParameter::saveProject()
{
    if (!m_hasLoaded || m_projectFilePath.isEmpty()) return false;

    QJsonObject reservoir;
    if(m_fullProjectData.contains("reservoir")) reservoir = m_fullProjectData["reservoir"].toObject();
    reservoir["porosity"] = m_phi;
    reservoir["thickness"] = m_h;
    reservoir["wellRadius"] = m_rw;
    reservoir["productionRate"] = m_q;
    m_fullProjectData["reservoir"] = reservoir;

    QJsonObject pvt;
    if(m_fullProjectData.contains("pvt")) pvt = m_fullProjectData["pvt"].toObject();
    pvt["viscosity"] = m_mu;
    pvt["volumeFactor"] = m_B;
    pvt["compressibility"] = m_Ct;
    m_fullProjectData["pvt"] = pvt;

    // 注意：保存 saveProject 时不写入 plotting_data 到 .pwt，保持分离
    // 为了防止 m_fullProjectData 中残留 plotting_data 被写入 .pwt，可以先移除
    QJsonObject dataToWrite = m_fullProjectData;
    dataToWrite.remove("plotting_data");

    QFile file(m_projectFilePath);
    if (!file.open(QIODevice::WriteOnly)) return false;
    file.write(QJsonDocument(dataToWrite).toJson());
    file.close();

    return true;
}

void ModelParameter::closeProject()
{
    m_hasLoaded = false;
    m_projectPath.clear();
    m_projectFilePath.clear();
    m_fullProjectData = QJsonObject();
    m_phi=0.05; m_h=20.0; m_mu=0.5; m_B=1.05; m_Ct=5e-4; m_q=50.0; m_rw=0.1;
}

void ModelParameter::saveFittingResult(const QJsonObject& fittingData)
{
    if (m_projectFilePath.isEmpty()) return;
    m_fullProjectData["fitting"] = fittingData;

    // 拟合结果依然保存在 .pwt 中（根据需求也可以分离，这里暂存主文件）
    QFile file(m_projectFilePath);
    if (file.open(QIODevice::WriteOnly)) {
        // 同样移除 plotting_data 再保存
        QJsonObject dataToWrite = m_fullProjectData;
        dataToWrite.remove("plotting_data");
        file.write(QJsonDocument(dataToWrite).toJson());
        file.close();
    }
}

QJsonObject ModelParameter::getFittingResult() const
{
    return m_fullProjectData.value("fitting").toObject();
}

// [修改] 保存绘图数据到单独的 json 文件
void ModelParameter::savePlottingData(const QJsonArray& plots)
{
    if (m_projectFilePath.isEmpty()) return;

    // 1. 更新内存缓存 (以便 getPlottingData 能立即获取)
    m_fullProjectData["plotting_data"] = plots;

    // 2. 构造独立文件路径
    QString dataFilePath = getPlottingDataFilePath();

    // 3. 构造要写入的对象
    QJsonObject dataObj;
    dataObj["plotting_data"] = plots;

    // 4. 写入文件
    QFile file(dataFilePath);
    if (file.open(QIODevice::WriteOnly)) {
        file.write(QJsonDocument(dataObj).toJson());
        file.close();
        qDebug() << "图表数据已保存至独立文件:" << dataFilePath;
    } else {
        qDebug() << "保存图表数据失败:" << dataFilePath;
    }
}

QJsonArray ModelParameter::getPlottingData() const
{
    return m_fullProjectData.value("plotting_data").toArray();
}
