## Conda
```bash
# create environment
conda create -n <name-env> python=3.8

# activate environment
conda activate <name-env>

# deactivate environment
conda deactivate
```

#### Recreate environment
```bash
conda env create -f environment.yml
```

#### Update environment
```bash
# save updated environment.yml
conda env export > environment.yml

# update environment with environment.yml
conda env update -f environment.yml
```

#### Conda shell
```bash
conda init (-all)
# then restart terminal

# exit conda shell
conda deactivate
```